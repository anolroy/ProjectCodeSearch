VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRetentionMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retentions"
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17775
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRetentionMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   17775
   Begin VB.PictureBox picBankCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   7830
      ScaleHeight     =   3180
      ScaleWidth      =   4035
      TabIndex        =   73
      Top             =   8775
      Visible         =   0   'False
      Width           =   4065
      Begin VB.CommandButton cmdBankClose 
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
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   90
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
         Height          =   2715
         Left            =   45
         TabIndex        =   75
         Top             =   450
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   4789
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
      Begin MSForms.Label Label9 
         Height          =   195
         Left            =   180
         TabIndex        =   77
         Top             =   135
         Width           =   1230
         VariousPropertyBits=   8388627
         Caption         =   "Bank Code"
         Size            =   "2170;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label8 
         Height          =   195
         Left            =   1725
         TabIndex        =   76
         Top             =   150
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Bank AC Name"
         Size            =   "2090;344"
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
         Index           =   8
         Left            =   90
         Top             =   90
         Width           =   3600
      End
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Automatic Demand Generate:"
      ForeColor       =   &H00FF00FF&
      Height          =   2220
      Left            =   6210
      TabIndex        =   58
      Top             =   5445
      Visible         =   0   'False
      Width           =   3715
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E5E5E5&
         Height          =   2100
         Index           =   0
         Left            =   40
         ScaleHeight     =   2040
         ScaleWidth      =   3555
         TabIndex        =   59
         Top             =   50
         Width           =   3615
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
            TabIndex        =   64
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtSearchToD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2025
            MaxLength       =   80
            TabIndex        =   63
            Top             =   1125
            Width           =   1380
         End
         Begin VB.TextBox txtSearchFromD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   80
            TabIndex        =   62
            Top             =   1125
            Width           =   1290
         End
         Begin VB.TextBox txtSearchRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   61
            Top             =   790
            Width           =   2685
         End
         Begin VB.TextBox txtSearchNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   10
            TabIndex        =   60
            Top             =   450
            Width           =   2685
         End
         Begin VB.CommandButton cmdSearchCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   2055
            TabIndex        =   67
            Top             =   1635
            Width           =   1200
         End
         Begin VB.CommandButton cmdSearchOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   65
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
            Top             =   450
            Width           =   225
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            Height          =   1155
            Index           =   7
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
            TabIndex        =   66
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
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   8055
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   43
      Top             =   9585
      Visible         =   0   'False
      Width           =   6285
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
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   45
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7091
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   51
         Top             =   375
         Width           =   4545
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8017;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   90
         TabIndex        =   50
         Top             =   375
         Width           =   1485
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2619;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1620
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   46
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   15
         Left            =   45
         Top             =   75
         Width           =   5850
      End
   End
   Begin VB.Frame fraAddTrans 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Automatic Demand Generate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1575
      Left            =   2640
      TabIndex        =   34
      Top             =   9615
      Visible         =   0   'False
      Width           =   3715
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E5E5E5&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   2
         Left            =   40
         ScaleHeight     =   1395
         ScaleWidth      =   3555
         TabIndex        =   35
         Top             =   40
         Width           =   3615
         Begin VB.CommandButton cmdCancelTrans 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   2100
            TabIndex        =   40
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddTrans 
            Caption         =   "&OK"
            Height          =   315
            Left            =   80
            TabIndex        =   39
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optBankReceipt 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Bank Receipt"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   405
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optBankTransfer 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Bank Transfer"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   37
            Top             =   735
            Width           =   1680
         End
         Begin VB.OptionButton optBankPayment 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Bank Payment"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   1335
            TabIndex        =   36
            Top             =   390
            Width           =   1695
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   2
            Height          =   660
            Index           =   4
            Left            =   75
            Top             =   360
            Width           =   3360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Add a Transaction:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   41
            Top             =   0
            Width           =   1410
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H0000C0C0&
            FillColor       =   &H00FF8080&
            FillStyle       =   0  'Solid
            Height          =   45
            Index           =   26
            Left            =   0
            Top             =   260
            Width           =   3855
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00808080&
            BorderColor     =   &H00C00000&
            BorderWidth     =   3
            Height          =   660
            Index           =   5
            Left            =   75
            Top             =   360
            Width           =   3360
         End
      End
   End
   Begin TabDlg.SSTab tabPayment 
      Height          =   9315
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   17685
      _ExtentX        =   31194
      _ExtentY        =   16431
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Retentions"
      TabPicture(0)   =   "frmRetentionMaster.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape4(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape4(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Shape4(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblBankRec(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblBankRec(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblBankRec(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblBankRec(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblBankRec(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblBankRec(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblBankRec(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblBankRec(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtClientList"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtPoperty"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(12)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkSelAll"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblBankRec(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblBankRec(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(13)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "fraButtons"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "flxRetention"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdClientList"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdProperty"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtDateFrom"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtDateTo"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdDisplay"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtRctTotal"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Retentions &History"
      TabPicture(1)   =   "frmRetentionMaster.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSearchHistory"
      Tab(1).Control(1)=   "cmdReverseHistory"
      Tab(1).Control(2)=   "cmdEdit"
      Tab(1).Control(3)=   "cmdClosebk(1)"
      Tab(1).Control(4)=   "flxRetentionHist"
      Tab(1).Control(5)=   "lblBankHist(8)"
      Tab(1).Control(6)=   "chkallHist"
      Tab(1).Control(7)=   "lblBankHist(4)"
      Tab(1).Control(8)=   "lblBankHist(6)"
      Tab(1).Control(9)=   "lblBankHist(7)"
      Tab(1).Control(10)=   "lblBankHist(5)"
      Tab(1).Control(11)=   "lblBankHist(3)"
      Tab(1).Control(12)=   "lblBankHist(2)"
      Tab(1).Control(13)=   "lblBankHist(1)"
      Tab(1).Control(14)=   "lblBankHist(0)"
      Tab(1).Control(15)=   "Shape4(0)"
      Tab(1).ControlCount=   16
      Begin VB.TextBox txtRctTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   16110
         MaxLength       =   80
         TabIndex        =   82
         Top             =   8820
         Width           =   1200
      End
      Begin VB.CommandButton cmdSearchHistory 
         Caption         =   "Sea&rch"
         Height          =   400
         Left            =   -72705
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   7110
         Width           =   1395
      End
      Begin VB.CommandButton cmdReverseHistory 
         Caption         =   "&Reverse History"
         Height          =   400
         Left            =   -74730
         TabIndex        =   55
         Top             =   7110
         Width           =   1935
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Display"
         Height          =   430
         Left            =   13410
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   450
         Width           =   960
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12285
         TabIndex        =   8
         Top             =   540
         Width           =   1065
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10575
         TabIndex        =   7
         Top             =   540
         Width           =   1020
      End
      Begin VB.CommandButton cmdProperty 
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
         Height          =   300
         Left            =   7830
         TabIndex        =   6
         Top             =   540
         Width           =   300
      End
      Begin VB.CommandButton cmdClientList 
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
         Height          =   300
         Left            =   4545
         TabIndex        =   5
         Top             =   540
         Width           =   300
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Transaction"
         Height          =   400
         Left            =   -70770
         TabIndex        =   24
         Top             =   7110
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdClosebk 
         Caption         =   "&Close"
         Height          =   400
         Index           =   1
         Left            =   -62250
         TabIndex        =   23
         Top             =   7080
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetention 
         Height          =   7395
         Left            =   2040
         TabIndex        =   11
         Top             =   1365
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   13044
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
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
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   17
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame fraButtons 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   430
            Left            =   180
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   1530
            Visible         =   0   'False
            Width           =   1450
         End
         Begin VB.CommandButton cmdNewBk 
            Caption         =   "&Add New"
            Height          =   430
            Left            =   135
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   180
            Width           =   1450
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Sea&rch"
            Height          =   420
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   5085
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdPostHistory 
            Caption         =   "Post to &History"
            Height          =   430
            Left            =   120
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3000
            Width           =   1450
         End
         Begin VB.CommandButton cmdPrintList 
            Caption         =   "&Print Retentions"
            Height          =   430
            Left            =   90
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   3600
            Width           =   1450
         End
         Begin VB.CommandButton cmdEditBk 
            Caption         =   "&Edit"
            Height          =   430
            Left            =   120
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   900
            Width           =   1450
         End
         Begin VB.CommandButton cmdClosebk 
            Cancel          =   -1  'True
            Caption         =   "C&lose"
            Height          =   430
            Index           =   0
            Left            =   120
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5880
            Width           =   1450
         End
         Begin VB.Shape Shape1 
            Height          =   2055
            Index           =   0
            Left            =   0
            Top             =   30
            Width           =   1695
         End
         Begin VB.Shape Shape2 
            Height          =   2055
            Left            =   0
            Top             =   2745
            Width           =   1695
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetentionHist 
         Height          =   6315
         Left            =   -74880
         TabIndex        =   25
         Top             =   705
         Width           =   14190
         _ExtentX        =   25030
         _ExtentY        =   11139
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
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
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   17
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Retentions Balance:"
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
         Left            =   14265
         TabIndex        =   83
         Top             =   8820
         Width           =   1830
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Code"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   6750
         TabIndex        =   81
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "SL No"
         Height          =   255
         Index           =   8
         Left            =   -68295
         TabIndex        =   80
         Top             =   450
         Width           =   2130
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   255
         Index           =   8
         Left            =   16155
         TabIndex        =   79
         Top             =   1125
         Width           =   975
      End
      Begin MSForms.CheckBox chkallHist 
         Height          =   255
         Left            =   -74865
         TabIndex        =   72
         Top             =   450
         Width           =   375
         VariousPropertyBits=   746588179
         BackColor       =   15781855
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "661;450"
         Value           =   "0"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkSelAll 
         Height          =   255
         Left            =   2025
         TabIndex        =   71
         Top             =   1080
         Width           =   375
         VariousPropertyBits=   746588179
         BackColor       =   15781855
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "661;450"
         Value           =   "0"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Date To:"
         Height          =   195
         Index           =   3
         Left            =   11655
         TabIndex        =   54
         Top             =   585
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Date From:"
         Height          =   195
         Index           =   12
         Left            =   9765
         TabIndex        =   53
         Top             =   585
         Width           =   765
      End
      Begin MSForms.TextBox txtPoperty 
         Height          =   285
         Left            =   5625
         TabIndex        =   52
         Tag             =   "ALL"
         Top             =   540
         Width           =   2205
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3889;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   2025
         TabIndex        =   42
         Tag             =   "ALL"
         Top             =   540
         Width           =   2565
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "4524;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Index           =   4
         Left            =   -66765
         TabIndex        =   33
         Top             =   450
         Width           =   495
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Index           =   6
         Left            =   -63525
         TabIndex        =   32
         Top             =   450
         Width           =   2430
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   255
         Index           =   7
         Left            =   -61995
         TabIndex        =   31
         Top             =   450
         Width           =   615
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   255
         Index           =   5
         Left            =   -64785
         TabIndex        =   30
         Top             =   450
         Width           =   735
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Statement ID"
         Height          =   255
         Index           =   3
         Left            =   -70140
         TabIndex        =   29
         Top             =   450
         Width           =   2130
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Property ID"
         Height          =   255
         Index           =   2
         Left            =   -71130
         TabIndex        =   28
         Top             =   450
         Width           =   1545
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
         Height          =   255
         Index           =   1
         Left            =   -72510
         TabIndex        =   27
         Top             =   450
         Width           =   1050
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Retention No"
         Height          =   255
         Index           =   0
         Left            =   -74250
         TabIndex        =   26
         Top             =   450
         Width           =   2055
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Retention No"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   2685
         TabIndex        =   22
         Top             =   1125
         Width           =   1140
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Property ID"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   21
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Statement ID"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   8010
         TabIndex        =   20
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   11970
         TabIndex        =   19
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Index           =   7
         Left            =   13860
         TabIndex        =   18
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Line No"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   9465
         TabIndex        =   17
         Top             =   1125
         Width           =   615
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   10740
         TabIndex        =   16
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   3870
         TabIndex        =   15
         Top             =   1125
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   660
         Index           =   3
         Left            =   240
         Top             =   360
         Width           =   17355
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   6
         Left            =   4935
         TabIndex        =   14
         Top             =   555
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   13
         Top             =   555
         Width           =   465
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   660
         Index           =   6
         Left            =   270
         Top             =   405
         Width           =   17265
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   0
         Left            =   -74835
         Top             =   435
         Width           =   14145
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   2040
         Top             =   1125
         Width           =   15510
      End
   End
End
Attribute VB_Name = "frmRetentionMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bSortingCol(8) As Boolean
Dim sTextBox   As Single
Dim iBankPayRow As Long
Dim colTransactionIDOtherPIGrid As String
Public UserSessionID As String
Public bEditMode As Boolean
Dim sText As String
Private Sub chkallHist_Click()
   Dim iRow As Integer
   
   If Not chkallHist.Value Then
      For iRow = 1 To flxRetentionHist.Rows - 1
         flxRetentionHist.TextMatrix(iRow, 0) = ""
      Next iRow
   Else
      For iRow = 1 To flxRetentionHist.Rows - 1
         flxRetentionHist.TextMatrix(iRow, 0) = "X"
      Next iRow
   End If
End Sub

Private Sub chkSelAll_Click()
    Dim iRow As Integer
   
   If Not chkSelAll.Value Then
      For iRow = 1 To flxRetention.Rows - 1
         flxRetention.TextMatrix(iRow, 0) = ""
      Next iRow
   Else
      For iRow = 1 To flxRetention.Rows - 1
         flxRetention.TextMatrix(iRow, 0) = "X"
      Next iRow
   End If
End Sub

Private Sub cmdaddPreview_Click()
    LoadForm frmRetentionAddPreview
End Sub

'Private Sub cmdBankAccountFiilter_Click()
'    sText = "1"
'    picBankCode.Visible = True
'    picBankCode.Top = cmdBankAccountFiilter.Top
'    picBankCode.Left = Label3(1).Left
'    Call LoadBankAccount
'End Sub

Private Sub cmdBankClose_Click()
    picBankCode.Visible = False
End Sub

Private Sub cmdDelete_Click()
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   Dim strID As String
   Dim strStatementID As String
   Dim adoconn As New ADODB.Connection
   For rCount = 1 To flxRetention.Rows - 1
        If flxRetention.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select a Retention.", vbInformation + vbOKOnly, "Statement Selection"
'      chkSelectAllDemands.Value = 0
      'ClearGridSelection
      Exit Sub
   End If
   strID = flxRetention.TextMatrix(selRow, 10)
   strStatementID = flxRetention.TextMatrix(selRow, 5)
   adoconn.Open getConnectionString
   If iIncDec = 1 Then
        If MsgBox("Do wish to delete this record?", vbYesNo, "Please confirm?") = vbYes Then
            If strStatementID <> "" Then
                 MsgBox "You cannot delete this Retention, because it is assigned to a Client Statement"
            Else
                 adoconn.Execute "Update RetentionDetails SET isDeleted=true  where ID=" & strID & ""
                 MsgBox "Retention record has been deleted"
            End If
         End If
    ElseIf iIncDec > 1 Then
             If MsgBox("Do wish to delete multiple records?", vbYesNo, "Please confirm?") = vbYes Then
                    For rCount = 1 To flxRetention.Rows - 1
                           If flxRetention.TextMatrix(rCount, 0) = "X" Then
                               iIncDec = iIncDec + 1
                               selRow = rCount
                               strID = flxRetention.TextMatrix(selRow, 10)
                               strStatementID = flxRetention.TextMatrix(selRow, 5)
                                If strStatementID <> "" Then
                                      ' MsgBox "You cannot delete this Retention, because it is assigned to a Client Statement", vbInformation, ""
                                  Else
                                       adoconn.Execute "Update RetentionDetails SET isDeleted=true  where ID=" & strID & ""
                                      
                                  End If
                           End If
                      Next
                    MsgBox "Retention records have been deleted", vbInformation + vbOKOnly, "Record delete"
              End If
    End If
            
  
   Call LoadflxRetention(adoconn, "")
   adoconn.Close
End Sub

Private Sub cmdSearchCancel_Click()
    Dim adoconn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        fraSearch.Visible = False
        adoconn.Open getConnectionString
        If cmdSearch.Caption = "Clear Sea&rch" Then
             cmdSearch.Caption = "Sea&rch"
        End If
        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
        Else
             Call LoadflxRetention(adoconn, "")
        End If
        adoconn.Close
    End If
End Sub

Private Sub cmdSearchHistory_Click()
    fraSearch.Left = 3015
    fraSearch.Top = 5085
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    If cmdSearchHistory.Caption = "Clear Sea&rch" Then
         txtSearchNo.text = ""
         txtSearchRef.text = ""
         cmdSearchHistory.Caption = "Sea&rch"
         fraSearch.Visible = False
    Else
        If fraSearch.Visible = False Then
            fraSearch.Visible = True
            txtSearchNo.SetFocus
        Else
            fraSearch.Visible = False
        End If
    End If
End Sub

Private Sub cmdSearchOK_Click()
    fraSearch.Visible = False
    Dim adoconn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        adoconn.Open getConnectionString
        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
            LoadflxRetention adoconn, ""
            'fmeLoading.Visible = False      cmdSearch.Caption = "Sea&rch"
        ElseIf Trim(txtSearchNo.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchRef.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
            LoadflxRetention adoconn, "3"
            cmdSearch.Caption = "Clear Sea&rch"
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
            cmdSearch.Caption = "Clear Sea&rch"
            If tabPayment.Tab = 0 Then
                LoadflxRetention adoconn, "3"
            Else
                LoadflxRetentionHist adoconn, "3"
            End If
        End If
        adoconn.Close
        Set adoconn = Nothing
    End If
End Sub

'Private Sub cmdTransTypeFilter_Click()
'    sText = "2"
'    picBankCode.Visible = True
'    picBankCode.Top = cmdTransTypeFilter.Top
'    picBankCode.Left = Label3(0).Left
'    Call LoadBankTranType
'End Sub
Private Sub LoadBankTranType()
    Dim rRow As Integer
    configGridBankTranType
    gridBankCode.Rows = 5
    rRow = 1
    gridBankCode.TextMatrix(rRow, 1) = "ALL"
    gridBankCode.TextMatrix(rRow, 2) = "ALL"
    
    rRow = 2
    gridBankCode.TextMatrix(rRow, 1) = "BP"
    gridBankCode.TextMatrix(rRow, 2) = "Bank Payment"
    
    rRow = 3
    gridBankCode.TextMatrix(rRow, 1) = "BR"
    gridBankCode.TextMatrix(rRow, 2) = "Bank Receipt"
    
    rRow = 4
    gridBankCode.TextMatrix(rRow, 1) = "BT"
    gridBankCode.TextMatrix(rRow, 2) = "Bank Transfer"
End Sub
'Private Sub gridBankCode_Click()
'    Dim adoconn As New ADODB.Connection
'    adoconn.Open getConnectionString
'    If sText = "1" Then
'        txtBankAccountFilter.text = gridBankCode.TextMatrix(gridBankCode.row, 1)
'        picBankCode.Visible = False
'        LoadflxRetention1 adoconn, 0, "ASC"
'    ElseIf sText = "2" Then
'        txtTransTypeFilter.Tag = gridBankCode.TextMatrix(gridBankCode.row, 1)
'        txtTransTypeFilter.text = gridBankCode.TextMatrix(gridBankCode.row, 2)
'        picBankCode.Visible = False
'        LoadflxRetention1 adoconn, 0, "ASC"
'    End If
'    adoconn.Close
'    Set adoconn = Nothing
'End Sub

Private Sub tabPayment_Click(PreviousTab As Integer)
     Dim adoconn As New ADODB.Connection
     If tabPayment.Tab = 1 Then
        adoconn.Open getConnectionString
        Call LoadflxRetentionHist(adoconn, "")
        adoconn.Close
      End If
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
        txtSearchToD.SetFocus
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

Private Sub txtSearchNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSearchRef.SetFocus
    End If
End Sub

Private Sub txtSearchRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
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
        cmdSearchOK.SetFocus
    End If
    TextBoxKeyPrsDate txtSearchToD, KeyAscii
End Sub

Private Sub txtSearchToD_LostFocus()
    If txtSearchToD.text <> "" Then TextBoxFormatDate txtSearchToD
End Sub
'Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTID;"
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
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Clients"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboC.Column() = Data()
'   cboC.ListIndex = 0
'   adoRst.Close
''*************************************** PROPERTY ******************************************
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "WHERE PropertyID <> '' " & _
'           "ORDER BY PropertyID;"
''   Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Properties"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'
'   cboP.Column() = Data()
'   cboP.ListIndex = 0
'
'NoRes:
'   adoRst.Close
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
'Private Sub LoadProperty(adoConn As ADODB.Connection)
'Dim szSQL As String
' Dim TotalRow As Integer, TotalCol As Integer
' 'Dim adoconn As New ADODB.Connection
' Dim adoRst As New ADODB.Recordset
'   Dim i As Integer, j As Integer
'   'adoconn.Open getConnectionString
'   If cmbClient.text = "ALL" Then
'            szSQL = "SELECT PropertyID, PropertyName, " & _
'                        "ProAddressLine1, ProPostCode " & _
'                    "FROM Property " & _
'                    "WHERE PropertyID <> '' " & _
'                    "ORDER BY PropertyID;"
'           Else
'                szSQL = "SELECT PropertyID, PropertyName, " & _
'                        "ProAddressLine1, ProPostCode " & _
'                    "FROM Property " & _
'                    "WHERE clientID = '" & txtClientList.tag & "' " & _
'                    "ORDER BY PropertyID;"
'           End If
''   Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Properties"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'
'   cmbProperty.Column() = Data()
'   cmbProperty.ListIndex = 0
'
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub
Private Sub cboBC_Change()
   SortTheGrid
End Sub

Private Sub cmbClient_Click()
  ' SortTheGrid
'   Dim adoconn As New ADODB.Connection
'   adoconn.Open getConnectionString
'   If cmbClient.text <> "ALL" Then
'        LoadflxRetentionbyclient adoconn
'   Else
'        LoadflxRetention adoconn
'   End If
'   loadProperty adoconn
'   adoconn.Close
End Sub

Private Sub SortTheGrid()
  ' If cmbProperty.ListCount <= 0 Then Exit Sub

   Dim i As Integer

   If txtClientList.Tag <> "ALL" Then
      For i = 1 To flxRetention.Rows - 1
         If flxRetention.TextMatrix(i, 13) = txtClientList.Tag Then
            flxRetention.RowHeight(i) = 240
         Else
            flxRetention.RowHeight(i) = 0
         End If
      Next i
   Else
      For i = 1 To flxRetention.Rows - 1
         flxRetention.RowHeight(i) = 240
      Next i
   End If

   If txtPoperty.Tag <> "ALL" Then
      For i = 1 To flxRetention.Rows - 1
         If flxRetention.TextMatrix(i, 11) = txtPoperty.Tag And flxRetention.RowHeight(i) = 240 Then
            flxRetention.RowHeight(i) = 240
         Else
            flxRetention.RowHeight(i) = 0
         End If
      Next i
   End If

'   If cboBC.Value <> "" Then
'      For i = 1 To flxRetention.Rows - 1
'         If flxRetention.TextMatrix(i, 14) = cboBC.Value And flxRetention.RowHeight(i) = 240 Then
'            flxRetention.RowHeight(i) = 240
'         Else
'            flxRetention.RowHeight(i) = 0
'         End If
'      Next i
'   End If
End Sub

Private Sub cmbProperty_Click()
   SortTheGrid
'    Dim adoconn As New ADODB.Connection
'   adoconn.Open getConnectionString
'   If cmbProperty.text <> "ALL" Then
'        LoadflxRetentionbyProperty adoconn
'   Else
'        LoadflxRetentionbyclient adoconn
'   End If
'   adoconn.Close
End Sub

Private Sub cmdClose_Click()

End Sub

Private Sub cmdCancelTrans_Click()
   tabPayment.Enabled = True
   fraAddTrans.Visible = False
End Sub

Private Sub cmdClientList_Click()
    picClient.Left = 1269.029
    picClient.Top = 455.299
    sTextBox = "1"
    LoadflxClient
    tabPayment.Enabled = False
   
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
   
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new

   
   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           flxClient.TextMatrix(1, 0) = ""
           flxClient.TextMatrix(1, 1) = "ALL"
           flxClient.TextMatrix(1, 2) = "All Client"
           'added by anol 20170208
           flxClient.RowHeight(1) = 280
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
               flxClient.RowHeight(rRow) = 280
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub
Private Sub cmdCloseBk_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub

'Private Sub cmdCopy_Click()
'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + fraButtons.Top + cmdCopy.Top + frmRetentionMaster.Top + tabPayment.Top + 1150
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + fraButtons.Left + frmRetentionMaster.Left + tabPayment.Left + cmdCopy.Left + 80
'   frmPopUpMenu.CallingFrom "BankCopy"
'   frmPopUpMenu.Show
'End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub cmdDisplay_Click()
    Exit Sub
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        
        
        If cmdDisplay.Caption = "&Display" And txtDateFrom.text <> "" And txtDateTo.text <> "" Then
             cmdDisplay.Caption = "&Clear"
        Else
            cmdDisplay.Caption = "&Display"
            txtDateFrom.text = ""
            txtDateTo.text = ""
        End If
        ConfigflxRetention
        LoadflxRetention1 adoconn, 0, "ASC"
        adoconn.Close
End Sub

Private Sub cmdEdit_Click()
   If flxRetentionHist.TextMatrix(flxRetentionHist.row, 0) = "" Then
      Exit Sub
   End If
   If flxRetentionHist.TextMatrix(flxRetentionHist.row, 15) = "Y" Then
      MsgBox "The transaction has been bank reconciled."
      Exit Sub
   End If
   Load frmBankTranEdit
   frmBankTranEdit.Caption = "Edit Bank Transaction"
   frmBankTranEdit.FrmBankTranEdit_CALLING_FROM = Me.Name & "History"

   With frmBankTranEdit
      .szTransID = flxRetentionHist.TextMatrix(flxRetentionHist.row, 11)
      .txtClientList.Tag = flxRetentionHist.TextMatrix(flxRetentionHist.row, 12)
      .txtClientList.text = flxRetentionHist.TextMatrix(flxRetentionHist.row, 18)
      .txtBankCode.text = flxRetentionHist.TextMatrix(flxRetentionHist.row, 1)
      .txtBankName.text = flxRetentionHist.TextMatrix(flxRetentionHist.row, 19)
      .txtProperty.Tag = flxRetentionHist.TextMatrix(flxRetentionHist.row, 13)
      .txtUnit.Tag = flxRetentionHist.TextMatrix(flxRetentionHist.row, 16)
     
      .txtDetails.text = flxRetentionHist.TextMatrix(flxRetentionHist.row, 7)
      .txtReference.text = flxRetentionHist.TextMatrix(flxRetentionHist.row, 6)
      .txtNC.Tag = flxRetentionHist.TextMatrix(flxRetentionHist.row, 4)
      .txtNC.text = flxRetentionHist.TextMatrix(flxRetentionHist.row, 20)
      .txtFund.Tag = flxRetentionHist.TextMatrix(flxRetentionHist.row, 14)
      .txtFund.Tag = flxRetentionHist.TextMatrix(flxRetentionHist.row, 6)
      .txtDate.text = flxRetentionHist.TextMatrix(flxRetentionHist.row, 2)
      .txtNet.text = Format(flxRetentionHist.TextMatrix(flxRetentionHist.row, 8), "0.00")
      '.cboVat.Value = flxRetentionHist.TextMatrix(flxRetentionHist.row, 17)
      .Label1(24).Caption = flxRetentionHist.TextMatrix(flxRetentionHist.row, 17)
      .txtVat_.text = Format(Val(flxRetentionHist.TextMatrix(flxRetentionHist.row, 9)), "0.00")
      .txtTotal.text = Format(Val(flxRetentionHist.TextMatrix(flxRetentionHist.row, 8)) + _
                       Val(flxRetentionHist.TextMatrix(flxRetentionHist.row, 9)), "0.00")
   End With

   frmBankTranEdit.Show
   Me.Enabled = False
End Sub

Private Sub cmdAddTrans_Click()
   If optBankReceipt.Value Or optBankPayment.Value Then
      Load frmBankTranEdit
      frmBankTranEdit.FrmBankTranEdit_CALLING_FROM = frmRetentionMaster.Name

      frmBankTranEdit.Caption = "Add " & IIf(optBankReceipt.Value, optBankReceipt.Caption, optBankPayment.Caption)

      frmBankTranEdit.Show
   End If
   If optBankTransfer.Value Then
      Load frmBankTransfer
      frmBankTransfer.FrmBankTransfer_CALLING_FROM = "frmRetentionMaster"

      frmBankTransfer.Show
   End If
   fraAddTrans.Visible = False
   tabPayment.Enabled = True
   frmRetentionMaster.Enabled = False
End Sub

Private Sub cmdEditBk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub cmdNewBk_Click()
'   fraAddTrans.Left = fraButtons.Left + cmdNewBk.Left + cmdNewBk.Width + tabPayment.Left
'   fraAddTrans.Top = fraButtons.Top + cmdNewBk.Top + tabPayment.Top
'   fraAddTrans.Visible = True
'   fraAddTrans.ZOrder 0
'
'   tabPayment.Enabled = False
'   cmdAddTrans.SetFocus
'   bEditMode = False
    LoadForm frmRetentionAdd

End Sub

Private Sub cmdPicCLose_Click()
        picClient.Visible = False
        tabPayment.Enabled = True
        cmdClientList.SetFocus
End Sub

'Private Sub cmdNewBk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   fraAddTrans.Left = fraButtons.Left + cmdNewBk.Left + cmdNewBk.Width + tabPayment.Left
'   fraAddTrans.Top = fraButtons.Top + cmdNewBk.Top + tabPayment.Top
'   fraAddTrans.Visible = True
'   fraAddTrans.ZOrder 0
'
''   tabPayment.Enabled = False
'   cmdAddTrans.SetFocus
'End Sub

Private Sub cmdPostHistory_Click()
'    Exit Sub
   If MsgBox("Do you wish to post these Retentions to history?", vbQuestion + vbYesNo, "Post to history") = vbNo Then Exit Sub

   Dim szBank As String
   Dim szSQL   As String
   Dim rCount As Integer
   Dim selRow As Integer
   Dim iIncDec As Integer
   Dim strID As String
   Dim strStatementID As String
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
       
   For rCount = 1 To flxRetention.Rows - 1
           If flxRetention.TextMatrix(rCount, 0) = "X" Then
               iIncDec = iIncDec + 1
               selRow = rCount
               strID = flxRetention.TextMatrix(selRow, 10)
               strStatementID = flxRetention.TextMatrix(selRow, 5)
                'If strStatementID <> "" Then
                      ' MsgBox "You cannot delete this Retention, because it is assigned to a Client Statement", vbInformation, ""
                'Else
                      adoconn.Execute "Update RetentionDetails SET isHist=true  where ID=" & strID & ""
                      
                'End If
           End If
    Next
    MsgBox "Retention have been posted to history successfully", vbInformation + vbOKOnly, "Record delete"
    Call LoadflxRetention(adoconn, "")
    adoconn.Close

End Sub

Private Function SelectedDemandID() As String
   Dim i As Integer

   SelectedDemandID = ""
   For i = 1 To flxRetention.Rows - 1
      If flxRetention.TextMatrix(i, 0) = "X" Then
         SelectedDemandID = SelectedDemandID & "'" & flxRetention.TextMatrix(i, 10) & "'"
         SelectedDemandID = SelectedDemandID & ","
      End If
   Next i
   If SelectedDemandID = "" Then Exit Function

   SelectedDemandID = Left(SelectedDemandID, Len(SelectedDemandID) - 1)
End Function

Private Sub cmdPrintList_Click()
'    Exit Sub

    Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    Dim strID As String
    Dim strStatementID As String
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    adoconn.Execute "Update RetentionDetails SET isPrint=false"
    For rCount = 1 To flxRetention.Rows - 1
         'If flxRetention.TextMatrix(rCount, 0) = "X" Then
            ' iIncDec = iIncDec + 1
             selRow = rCount
             strID = flxRetention.TextMatrix(selRow, 10)
             strStatementID = flxRetention.TextMatrix(selRow, 5)
             adoconn.Execute "Update RetentionDetails SET isPrint=true  where ID=" & strID & ""
        ' End If
    Next

    adoconn.Close


   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RententionDetailsListingReport.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdproperty_Click()
    picClient.Left = 5445.029
    picClient.Top = 455.299
    sTextBox = "2"
    LoadPropertyList
    tabPayment.Enabled = False
   
    picClient.Visible = True
    txtSearchClientID.SetFocus
    
End Sub

Private Sub cmdReverseHistory_Click()
    'Added by anol 24 May 2016

   Dim szSQL As String
   Dim iRow As Integer, szPurID As String
   Dim adoconn As New ADODB.Connection
   Dim strPartSql As String
   Dim rCount As Integer
   Dim iIncDec As Integer
   Dim selRow As Integer
   Dim strID As String
   Dim strStatementID As String
   strPartSql = SelPurHistory
   On Error GoTo Catch_Error
   If strPartSql = "" Then Exit Sub
   If MsgBox("Are you sure to reverse back selected retention from the history?", vbQuestion + vbYesNo, "reverse History") = vbNo Then Exit Sub
   'szPurID = SelPurHistory()

   adoconn.Open getConnectionString
   
    For rCount = 1 To flxRetentionHist.Rows - 1
        If flxRetentionHist.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
            strID = flxRetentionHist.TextMatrix(selRow, 10)
            strStatementID = flxRetentionHist.TextMatrix(selRow, 5)
            adoconn.Execute "Update RetentionDetails SET isHist=False  where ID=" & strID & ""
        End If
      Next

   Call LoadflxRetention(adoconn, "")
   Call LoadflxRetentionHist(adoconn, "")

   adoconn.Close
   Set adoconn = Nothing
   Exit Sub

Catch_Error:
   MsgBox "Select a Retention to reverese.", vbCritical + vbOKOnly, "Warning"
End Sub
Private Function SelPurHistory() As String
   Dim i As Integer

   SelPurHistory = ""
   For i = 1 To flxRetentionHist.Rows - 1
      If flxRetentionHist.TextMatrix(i, 0) = "X" Then
         'SelPurHistory = SelPurHistory & CStr(Mid(flxRetentionHist.TextMatrix(i, 1), 3, Len(flxRetentionHist.TextMatrix(i, 1))))
         SelPurHistory = SelPurHistory & "'" & CStr(flxRetentionHist.TextMatrix(i, 12)) & "'"
         SelPurHistory = SelPurHistory & ","
      End If
   Next i
   If Len(SelPurHistory) = 0 Then Exit Function
   SelPurHistory = Left(SelPurHistory, Len(SelPurHistory) - 1)
End Function

Private Sub cmdSearch_Click()
    fraSearch.Left = 3015
    fraSearch.Top = 5085
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    txtSearchNo.text = ""
'    txtSearchRef.text = ""
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    If cmdSearch.Caption = "Clear Sea&rch" Then
         'Call LoadflxRetention(adoconn, "")
         txtSearchNo.text = ""
         txtSearchRef.text = ""
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
    
'    adoConn.Close
'    Set adoConn = Nothing
End Sub

Private Sub flxRetention_Click()
   If flxRetention.TextMatrix(flxRetention.row, 0) = "" And flxRetention.TextMatrix(flxRetention.row, 1) <> "" Then
      flxRetention.TextMatrix(flxRetention.row, 0) = "X"
   Else
      flxRetention.TextMatrix(flxRetention.row, 0) = ""
   End If
'    Dim szSlNo As String
'        Dim iIncDec As Integer
'        If flxRetention.TextMatrix(flxRetention.row, 2) = "" Then Exit Sub
''        If flxRetention.col = 0 Then
''            iIncDec = iIncDec + SelectFlxGridRow(0, flxRetention, flxRetention.row) 'Returns 1 or -1 depends on selection
''        End If
'            SelectOnly1RowFlxGrid flxRetention, flxRetention.row, 0
''
''        'SelectOnly1RowFlxGrid flxRetention, flxRetention.row, 0
''        If flxRetention.TextMatrix(flxRetention.row, 0) = "X" Then
''            If flxRetention.TextMatrix(flxRetention.row, 1) = "+" Then
''                szCurrentStatementID = Replace(flxRetention.TextMatrix(flxRetention.row, 2), "CS", "")
''               ' szAvailableFund1 = flxRetention.TextMatrix(flxRetention.row, 8)
''                Call LoadRentSummaryDetails
''            End If
''        End If
'
'        Call SquezeExpand
End Sub
Private Sub SquezeExpand()
       Dim i As Integer, iCurRowHeight As Integer

  iCurRowHeight = 280
   

   If flxRetention.col = 1 And flxRetention.TextMatrix(flxRetention.row, 1) = "+" Then          'Expanding the grid
      flxRetention.TextMatrix(flxRetention.row, 1) = ">"
      iCurRowHeight = flxRetention.RowHeight(flxRetention.row)
      i = 1

      While flxRetention.TextMatrix(flxRetention.row + i, 1) = "-"
         flxRetention.RowHeight(flxRetention.row + i) = iCurRowHeight
         i = i + 1
         If (flxRetention.row + i) = flxRetention.Rows Then Exit Sub
      Wend
      Exit Sub
   End If

   If flxRetention.col = 1 And flxRetention.TextMatrix(flxRetention.row, 1) = ">" Then          'Squeezing the grid
      flxRetention.TextMatrix(flxRetention.row, 1) = "+"
      i = 1
      While flxRetention.TextMatrix(flxRetention.row + i, 1) = "-"
         flxRetention.RowHeight(flxRetention.row + i) = 0
         i = i + 1
         If (flxRetention.row + i) = flxRetention.Rows Then Exit Sub
      Wend
      Exit Sub
   End If
   'HighLightRowFlxGridA flxRetention, flxRetention.row
End Sub
Public Sub CopyTransaction()
   Dim szStr As String
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

   On Error GoTo ErrHandler
'      connect to database
   adoconn.Open getConnectionString

   szStr = "SELECT BP.* " & _
           "FROM tlbBankPayment AS BP;"
'Debug.Print szStr
   adoRst.Open szStr, adoconn, adOpenDynamic, adLockOptimistic
   With adoRst
      If .EOF Then
         GoTo ErrHandler
      Else
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = flxRetention.TextMatrix(flxRetention.row, 13)
         .Fields.Item("BANK_AC").Value = flxRetention.TextMatrix(flxRetention.row, 14)
         .Fields.Item("PropertyID").Value = flxRetention.TextMatrix(flxRetention.row, 11)
         .Fields.Item("UNIT_ID").Value = flxRetention.TextMatrix(flxRetention.row, 12)
         .Fields.Item("DESCRIPTION").Value = flxRetention.TextMatrix(flxRetention.row, 8)
         .Fields.Item("PROJ_REF").Value = flxRetention.TextMatrix(flxRetention.row, 7)
         .Fields.Item("NOMINAL_CODE").Value = flxRetention.TextMatrix(flxRetention.row, 5)
         .Fields.Item("DEPT_ID").Value = flxRetention.TextMatrix(flxRetention.row, 6)
         .Fields.Item("TRAN_DATE").Value = Format(Now, "dd mmmm yyyy")
         .Fields.Item("NET_AMOUNT").Value = CCur(flxRetention.TextMatrix(flxRetention.row, 15))
         .Fields.Item("TAX_CODE").Value = flxRetention.TextMatrix(flxRetention.row, 18)
         .Fields.Item("VAT").Value = flxRetention.TextMatrix(flxRetention.row, 16)

         If Left(flxRetention.TextMatrix(flxRetention.row, 1), 2) = "BR" Then
            .Fields.Item("TransactionType").Value = 12
            .Fields.Item("TRANS").Value = "BR"
         End If
         If Left(flxRetention.TextMatrix(flxRetention.row, 1), 2) = "BP" Then
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
         End If

         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoconn)
         .Update
         .Close
      End If
   End With

   ShowMsgInTaskBar "The Transaction has been copied sucessfully.", "Y", "P"

   Set adoRst = Nothing

   flxRetention.Clear
   flxRetention.Rows = 2

   Call LoadflxRetention(adoconn, "")

   adoconn.Close
   Set adoconn = Nothing
   Exit Sub

ErrHandler:
   MsgBox "System could not update the record.", vbExclamation + vbOKOnly, "Edit Bank Transactions"

   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
End Sub

Public Sub CopyRevTransaction()
   Dim szStr As String
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

   On Error GoTo ErrHandler
'      connect to database
   adoconn.Open getConnectionString

   szStr = "SELECT BP.* " & _
           "FROM tlbBankPayment AS BP;"
'Debug.Print szStr
   adoRst.Open szStr, adoconn, adOpenDynamic, adLockOptimistic
   With adoRst
      If .EOF Then
         GoTo ErrHandler
      Else
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = flxRetention.TextMatrix(flxRetention.row, 13)
         .Fields.Item("BANK_AC").Value = flxRetention.TextMatrix(flxRetention.row, 14)
         .Fields.Item("PropertyID").Value = flxRetention.TextMatrix(flxRetention.row, 11)
         .Fields.Item("UNIT_ID").Value = flxRetention.TextMatrix(flxRetention.row, 12)
         .Fields.Item("DESCRIPTION").Value = flxRetention.TextMatrix(flxRetention.row, 8)
         .Fields.Item("PROJ_REF").Value = flxRetention.TextMatrix(flxRetention.row, 7)
         .Fields.Item("NOMINAL_CODE").Value = flxRetention.TextMatrix(flxRetention.row, 5)
         .Fields.Item("DEPT_ID").Value = flxRetention.TextMatrix(flxRetention.row, 6)
         .Fields.Item("TRAN_DATE").Value = Format(Now, "dd mmmm yyyy")
         .Fields.Item("NET_AMOUNT").Value = CCur(flxRetention.TextMatrix(flxRetention.row, 15))
         .Fields.Item("TAX_CODE").Value = flxRetention.TextMatrix(flxRetention.row, 18)
         .Fields.Item("VAT").Value = flxRetention.TextMatrix(flxRetention.row, 16)

         If Left(flxRetention.TextMatrix(flxRetention.row, 1), 2) = "BR" Then
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
         End If
         If Left(flxRetention.TextMatrix(flxRetention.row, 1), 2) = "BP" Then
            .Fields.Item("TransactionType").Value = 12
            .Fields.Item("TRANS").Value = "BR"
         End If

         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoconn)
         .Update
         .Close
      End If
   End With

   ShowMsgInTaskBar "The Transaction has been copy reveresed sucessfully.", "Y", "P"

   Set adoRst = Nothing

   flxRetention.Clear
   flxRetention.Rows = 2
      
   Call LoadflxRetention(adoconn, "")

   adoconn.Close
   Set adoconn = Nothing
   Exit Sub

ErrHandler:
   MsgBox "System could not update the record.", vbExclamation + vbOKOnly, "Edit Bank Transactions"

   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub flxRetention_DblClick()
    Exit Sub
    cmdEditBk_Click
End Sub
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
   
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new

   
        adoconn.Open getConnectionString
               If txtClientList.Tag <> "ALL" Then
                     szSQL = "SELECT PropertyID, PropertyName " & _
                       "FROM Property " & _
                       "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                       "ORDER BY PropertyID;"
               Else
                     szSQL = "SELECT PropertyID, PropertyName " & _
                       "FROM Property " & _
                       "ORDER BY PropertyID;"
               End If
               
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           flxClient.TextMatrix(1, 0) = ""
           flxClient.TextMatrix(1, 1) = "ALL"
           flxClient.TextMatrix(1, 2) = "All Property"
           flxClient.RowHeight(1) = 280
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
               flxClient.RowHeight(rRow) = 280
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub flxRetention_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow

   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub flxRetentionHistHist_Click()
    If flxRetentionHist.TextMatrix(flxRetentionHist.row, 0) = "" And flxRetentionHist.TextMatrix(flxRetentionHist.row, 1) <> "" Then
      flxRetentionHist.TextMatrix(flxRetentionHist.row, 0) = "X"
   Else
      flxRetentionHist.TextMatrix(flxRetentionHist.row, 0) = ""
   End If
End Sub

Private Sub flxRetentionHist_Click()
     If flxRetentionHist.TextMatrix(flxRetentionHist.row, 0) = "" And flxRetentionHist.TextMatrix(flxRetentionHist.row, 1) <> "" Then
      flxRetentionHist.TextMatrix(flxRetentionHist.row, 0) = "X"
   Else
      flxRetentionHist.TextMatrix(flxRetentionHist.row, 0) = ""
   End If
End Sub

'Private Sub flxRetentionHist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Me.MousePointer = vbArrow
'End Sub

Private Sub flxClient_Click()
    tabPayment.Enabled = True
    Dim adoconn As New ADODB.Connection
     adoconn.Open getConnectionString
    If sTextBox = "1" Then
            txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
            txtPoperty.Tag = "ALL"
            txtPoperty.text = "ALL"
            ConfigflxRetention
            LoadflxRetention1 adoconn, 0, "ASC"
            cmdProperty.SetFocus
    Else
            txtPoperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtPoperty.text = flxClient.TextMatrix(flxClient.row, 2)
            ConfigflxRetention
            LoadflxRetention1 adoconn, 0, "ASC"
            txtDateFrom.SetFocus
    End If
    adoconn.Close
    'SortTheGrid
    picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
         picClient.Visible = False
          tabPayment.Enabled = True
          
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           End If
    End If
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

Private Sub Form_Load()
   Me.Height = 9855
   Me.Width = 17865
'   tabPayment.Tab = 0
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   UserSessionID = GetTimeStamp
   Dim adoconn As New ADODB.Connection
   tabPayment.Tab = 0
   adoconn.Open getConnectionString
   Call ConfigflxRetention
   Call LoadflxRetention(adoconn, "")
   'ConfigflxRetentionHist
'   Call LoadflxRetentionHist(adoconn, "")
   'PrepareList adoConn, cmbClient, cmbProperty
   'LoadBankAccountInCombo adoConn
   
'   txtTransTypeFilter.Tag = "ALL"
'   txtTransTypeFilter.text = "ALL"
   adoconn.Close
   Set adoconn = Nothing

   Call WheelHook(Me.hWnd)
End Sub
Private Sub LoadBankAccount()
   ' Error Handler
   'On Error GoTo Error_Handler
   configGridBankCode

   Dim adoconn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoconn = New ADODB.Connection
   If txtClientList.text = "ALL" Then
        rRow = 1
        gridBankCode.TextMatrix(rRow, 1) = "ALL"
        gridBankCode.TextMatrix(rRow, 2) = "ALL"
        Exit Sub
   End If
         
   adoconn.Open getConnectionString

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = '" & txtClientList.Tag & "' AND " & _
               "NominalLedger.ClientID = '" & txtClientList.Tag & "';"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "Please setup bank account for this client : '" & txtClientList.text & "'", vbInformation, "Global bank account"
      picBankCode.Visible = False
   Else
      gridBankCode.Rows = adoRst.RecordCount + 2
      rRow = 1
        gridBankCode.TextMatrix(rRow, 1) = "ALL"
        gridBankCode.TextMatrix(rRow, 2) = "ALL"
      rRow = 2
      While Not adoRst.EOF
         gridBankCode.TextMatrix(rRow, 1) = adoRst.Fields.Item("BNC").Value
         gridBankCode.TextMatrix(rRow, 2) = adoRst.Fields.Item("BNN").Value
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
       picBankCode.Visible = True
       gridBankCode.row = 1
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoconn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "Prestige Database Error: ", vbExclamation, "Load Bank Account in Demand"

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoconn = Nothing
End Sub
Private Sub configGridBankCode()
   gridBankCode.Visible = True
   gridBankCode.Clear
   gridBankCode.Cols = 4
   gridBankCode.TextMatrix(0, 0) = ""
   gridBankCode.TextMatrix(0, 1) = ""
   gridBankCode.ColWidth(0) = 60
   gridBankCode.ColWidth(1) = 1200
   gridBankCode.ColAlignment(1) = vbLeftJustify
   gridBankCode.ColWidth(2) = 2600
   gridBankCode.ColAlignment(2) = vbLeftJustify
   gridBankCode.ColWidth(3) = 0
   gridBankCode.RowHeight(0) = 0
   gridBankCode.Rows = 2
   Label9.Caption = "Bank Code"
   Label8.Caption = "Bank Name"
   
End Sub
Private Sub configGridBankTranType()
   gridBankCode.Visible = True
   gridBankCode.Clear
   gridBankCode.Cols = 4
   gridBankCode.TextMatrix(0, 0) = ""
   gridBankCode.TextMatrix(0, 1) = ""
   gridBankCode.ColWidth(0) = 60
   gridBankCode.ColWidth(1) = 1200
   gridBankCode.ColAlignment(1) = vbLeftJustify
   gridBankCode.ColWidth(2) = 2600
   gridBankCode.ColAlignment(2) = vbLeftJustify
   gridBankCode.ColWidth(3) = 0
   gridBankCode.RowHeight(0) = 0
   gridBankCode.Rows = 2
   Label9.Caption = "Type"
   Label8.Caption = "Transactions"
   
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
     UnLoadForm Me
   If BANK_PAYMENT_HISTORY_LOADED Then
      BANK_PAYMENT_HISTORY_LOADED = False
      Unload frmBankPaymentHistory
   End If
End Sub

'Public Sub LoadflxRetentionHist(adoConn As ADODB.Connection, Filter As String)
'   Dim iRow As Integer
'   Dim szStr As String
'   Dim tempstr As String
'   Dim adoRst As New ADODB.Recordset
'
''   szStr = "SELECT BP.*, F.FundName, P.PropertyName,C.ClientName " & _
''           "FROM ((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = CSTR(F.FundID)) INNER JOIN CLIENT C ON C.ClientID=BP.ClientID)  LEFT JOIN " & _
''                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
''           "WHERE BP.DEPT_ID = CSTR(F.FundID) AND " & _
''               "BP.UPDATE_SAGE = TRUE " & _
''           "ORDER BY BP.TRANS, CLNG(BP.TRAN_ID);"
''  szStr = "SELECT BP.*, F.FundName, P.PropertyName,C.ClientName,NL.Name,(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1 " & _
''           "FROM (((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = CSTR(F.FundID))  " & _
''           "INNER JOIN CLIENT C ON C.ClientID=BP.ClientID) INNER JOIN NominalLedger NL ON NL.ClientID=BP.ClientID AND NL.Code=BP.NOMINAL_CODE) LEFT JOIN " & _
''                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
''           "WHERE BP.DEPT_ID = CSTR(F.FundID) AND " & _
''               "BP.UPDATE_SAGE = TRUE " & _
''           "ORDER BY BP.TRANS, CLNG(BP.TRAN_ID);" 'UPDATE_SAGE = TRUE means history is true
'    szStr = "SELECT TRANS & TRAN_ID as INVNO, BP.*, F.FundName, P.PropertyName,C.ClientName,NL.Name,(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1 " & _
'           "FROM (((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = CSTR(F.FundID))  " & _
'           "INNER JOIN CLIENT C ON C.ClientID=BP.ClientID) INNER JOIN NominalLedger NL ON NL.ClientID=BP.ClientID AND NL.Code=BP.NOMINAL_CODE) LEFT JOIN " & _
'                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
'           "WHERE BP.DEPT_ID = CSTR(F.FundID) AND " & _
'               "BP.UPDATE_SAGE = TRUE " & _
'            "ORDER BY  TRANS,CLNG(TRAN_ID) desc" 'UPDATE_SAGE = TRUE means history is true
'    If Filter = "3" Then
'        szStr = "SELECT TRANS & TRAN_ID as INVNO, BP.*, F.FundName, P.PropertyName,C.ClientName,NL.Name,(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1 " & _
'           "FROM (((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = CSTR(F.FundID))  " & _
'           "INNER JOIN CLIENT C ON C.ClientID=BP.ClientID) INNER JOIN NominalLedger NL ON NL.ClientID=BP.ClientID AND NL.Code=BP.NOMINAL_CODE) LEFT JOIN " & _
'                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
'           "WHERE BP.DEPT_ID = CSTR(F.FundID) AND " & _
'               "BP.UPDATE_SAGE = TRUE AND BP.TRAN_DATE >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND " & _
'                    "BP.TRAN_DATE <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "# " & _
'            "ORDER BY  TRANS,CLNG(TRAN_ID) desc"
'    End If
'
'   adoRst.Open szStr, adoConn, adOpenDynamic, adLockPessimistic
''Debug.Print szStr
''  szHeader$ = "<No|<Bank|<Date|<Property|<NC|<Ref|<Fund|<Details|>Net|>VAT|>Total|ID|Client|PropID|FundID|Recon|Unit|TC"
'
'    If Filter = "1" Then
'        If txtSearchNo.text <> "" Then
'            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
'            adoRst.Filter = "INvNo Like '%" & tempstr & "%'"
'        End If
'    End If
'
'
'
'
'    If Filter = "2" Then
'         If txtSearchRef.text <> "" Then
'            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
'            If tabPayment.Tab = 0 Then
'                adoRst.Filter = "Details Like '%" & tempstr & "%'"
'            Else
'                adoRst.Filter = "Description Like '%" & tempstr & "%'"
'            End If
'        End If
'    End If
'   flxRetentionHist.Clear
'   flxRetentionHist.Rows = 2
'   iRow = 1
'   While Not adoRst.EOF
'      flxRetentionHist.TextMatrix(iRow, 0) = ""
'      flxRetentionHist.TextMatrix(iRow, 1) = adoRst!InvNo 'adoRst!TRANS & adoRst!TRAN_ID
'      flxRetentionHist.TextMatrix(iRow, 2) = adoRst!BANK_AC
'      flxRetentionHist.TextMatrix(iRow, 3) = adoRst!TRAN_DATE
'      flxRetentionHist.TextMatrix(iRow, 4) = IIf(IsNull(adoRst!PropertyName), "", adoRst!PropertyName)
'      flxRetentionHist.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!Nominal_code), "", adoRst!Nominal_code)
'      flxRetentionHist.TextMatrix(iRow, 6) = IIf(IsNull(adoRst!PROJ_REF), "", adoRst!PROJ_REF)
'      flxRetentionHist.TextMatrix(iRow, 7) = IIf(IsNull(adoRst!FundName), "", adoRst!FundName)
'      flxRetentionHist.TextMatrix(iRow, 8) = IIf(IsNull(adoRst!description), "", adoRst!description)
'      flxRetentionHist.TextMatrix(iRow, 9) = Format(IIf(IsNull(adoRst!NET_AMOUNT), "", adoRst!NET_AMOUNT), "0.00")
'      flxRetentionHist.TextMatrix(iRow, 10) = Format(IIf(IsNull(adoRst!vat), "0", adoRst!vat), "0.00")
'      flxRetentionHist.TextMatrix(iRow, 11) = Format(Val(flxRetentionHist.TextMatrix(iRow, 8)) + _
'                                          Val(flxRetentionHist.TextMatrix(iRow, 9)), "0.00")
'      flxRetentionHist.TextMatrix(iRow, 12) = adoRst!My_ID
'      flxRetentionHist.TextMatrix(iRow, 13) = IIf(IsNull(adoRst!ClientID), "", adoRst!ClientID)
'      flxRetentionHist.TextMatrix(iRow, 14) = IIf(IsNull(adoRst!propertyID), "", adoRst!propertyID)
'      flxRetentionHist.TextMatrix(iRow, 15) = adoRst!DEPT_ID
'      flxRetentionHist.TextMatrix(iRow, 16) = IIf(IsNull(adoRst!ReconNow), "N", "Y")
'      flxRetentionHist.TextMatrix(iRow, 17) = IIf(IsNull(adoRst!UNIT_ID), "", adoRst!UNIT_ID)
'      flxRetentionHist.TextMatrix(iRow, 18) = IIf(IsNull(adoRst!TAX_CODE), "", adoRst!TAX_CODE)
'      flxRetentionHist.TextMatrix(iRow, 19) = IIf(IsNull(adoRst!ClientName), "", adoRst!ClientName)
'      flxRetentionHist.TextMatrix(iRow, 20) = IIf(IsNull(adoRst!Name), "", adoRst!Name) 'Nominal Bank Account Name
'       flxRetentionHist.TextMatrix(iRow, 21) = IIf(IsNull(adoRst!Name1), "", adoRst!Name1) 'Nominal Account Name
'      adoRst.MoveNext
'      If Not adoRst.EOF Then flxRetentionHist.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub ConfigflxRetentionHist()
  Dim szHeader As String, iCol As Integer

   flxRetentionHist.Clear
   flxRetentionHist.Rows = 2
   'flxRetentionHist.Cols = 25
   'adding 4 more col
   flxRetentionHist.Cols = 14
   flxRetentionHist.RowHeight(0) = 0

   szHeader$ = "ID|<Retention No|<ClientID|<Property ID|<Statement ID|<SL No|<Date|<Reference|<Description|<Amount"
   flxRetentionHist.FormatString = szHeader$

   flxRetentionHist.ColWidth(0) = 240
   iCol = 1
   flxRetentionHist.ColWidth(1) = 1500 ' lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left + 200
   'flxRetentionHist.ColWidth(2) = 1500  'lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left + 200
   For iCol = 2 To 7 'flxRetentionHist.Cols - 3 when cols was 10
      flxRetentionHist.ColWidth(iCol) = lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left
   Next iCol
    flxRetentionHist.ColWidth(iCol) = 1200 '8
    flxRetentionHist.ColAlignment(iCol) = vbRightJustify
    flxRetentionHist.ColWidth(iCol + 1) = 1200 '9 ID
    flxRetentionHist.ColAlignment(iCol + 1) = vbRightJustify
    flxRetentionHist.ColWidth(iCol + 2) = 0 'ID
    flxRetentionHist.ColWidth(iCol + 3) = 0 'Client Name
    flxRetentionHist.ColWidth(iCol + 4) = 0 'FundID
    flxRetentionHist.ColWidth(iCol + 5) = 0 'FundName
End Sub

Private Sub ConfigflxRetention()
   Dim szHeader As String, iCol As Integer

   flxRetention.Clear
   flxRetention.Rows = 2
   'flxRetention.Cols = 25
   'adding 4 more col
   flxRetention.Cols = 16
   flxRetention.RowHeight(0) = 0

   szHeader$ = "ID|<Retention No|<ClientID|<Property ID|<BankCode|<BankName|<Statement ID|<SL No|<Date|<Reference|<Description|<Amount"
   flxRetention.FormatString = szHeader$

   flxRetention.ColWidth(0) = 240
   iCol = 1
   flxRetention.ColWidth(1) = 1500 ' lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left + 200
   'flxRetention.ColWidth(2) = 1500  'lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left + 200
'   For iCol = 2 To 7 'flxRetention.Cols - 3 when cols was 10
'      flxRetention.ColWidth(iCol) = lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left
'      Debug.Print iCol
'      Debug.Print lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left
'   Next iCol
    flxRetention.ColWidth(2) = 1650 'ClientID
    flxRetention.ColWidth(3) = 1185 'Property ID
    flxRetention.ColWidth(4) = 1185 'Bank Code
    flxRetention.ColAlignment(4) = vbRightJustify
    flxRetention.ColWidth(5) = 0 'Bank Name
    flxRetention.ColWidth(6) = 1475  ' statement ID
    flxRetention.ColWidth(7) = 1050
    flxRetention.ColWidth(8) = 1365
    flxRetention.ColWidth(9) = 1820
    iCol = 10
    flxRetention.ColWidth(iCol) = 2600 '8
    flxRetention.ColAlignment(iCol) = vbRightJustify
    flxRetention.ColWidth(iCol + 1) = 1200 '9 ID
    flxRetention.ColAlignment(iCol + 1) = vbRightJustify
    flxRetention.ColWidth(iCol + 2) = 0 'ID
    flxRetention.ColWidth(iCol + 3) = 0 'Client Name
    flxRetention.ColWidth(iCol + 4) = 0 'FundID
    flxRetention.ColWidth(iCol + 5) = 0 'FundName
   
   
   
End Sub
Private Function InstantLockingCheck() As Boolean 'unlocking for all row
   Dim adoPay As New ADODB.Recordset
   Dim rsLockDialog As New ADODB.Recordset
   Dim adoconn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer
   Dim selRow As Integer
   Dim selcol As Integer
   Dim i As Integer
   Dim j As Integer
   Dim strSQL As String
   selRow = flxRetention.row
   selcol = flxRetention.col
   
   
   adoconn.Open getConnectionString
   
'   ' I am doing some test here
'   ' on loading time full table vs selected row
'
'    szSQL = " SELECT BP.TransactionID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment BP WHERE MYid= '" & flxRetention.TextMatrix(flxRetention.row, 10) & "'"
'
'   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   adoPay.Close
'   szSQL = " SELECT BP.TransactionID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment"
'
'   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   adoPay.Close
'
'   Exit Sub
   'first part is instant lock
   szSQL = " SELECT BP.TransactionID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment BP WHERE MYid= '" & flxRetention.TextMatrix(flxRetention.row, 10) & "'"
   
   adoPay.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

  'locking status show for current row
   If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 Then   'szSQL <> UserSessionID shall be always true bcoz PI is generating only one session thrgh out this module.
                flxRetention.col = 0
                flxRetention.row = flxRetention.row
                flxRetention.CellBackColor = vbRed
                InstantLockingCheck = False
                colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoPay("TransactionID").Value), "", adoPay("TransactionID").Value) & ","
            Else 'lock for this user
                        flxRetention.col = 0
                        i = flxRetention.row
                        flxRetention.CellBackColor = vbWhite
    '                    adoconn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Invoice',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
    '                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where tlbPayment.PI = '" & flxRetention.TextMatrix(iPIEdit, 0) & "'"
                        'Need to clear the locking flag
                        flxRetention.TextMatrix(i, 25) = ""
                        flxRetention.TextMatrix(i, 26) = ""
                        flxRetention.TextMatrix(i, 27) = ""
                        flxRetention.TextMatrix(i, 28) = ""
            End If
           
   End If
   'second part instant unlock
       If Len(colTransactionIDOtherPIGrid) > 0 Then
            szSQL = "SELECT TransactionID,UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
                 "FROM tlbBankPayment where (isnull(UserSessionID) OR UserSessionID='') " & _
                 " AND TransactionID in (" & colTransactionIDOtherPIGrid & ") order by 2,3 Desc;"
            rsLockDialog.Open szSQL, adoconn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
             While Not rsLockDialog.EOF
                      flxRetention.col = 0
                      For j = 1 To flxRetention.Rows - 1
                          If flxRetention.TextMatrix(j, 10) = rsLockDialog("transactionID").Value And i <> j Then 'no need to update row of first part check
                                flxRetention.row = j
                                flxRetention.CellBackColor = vbWhite
                          End If
                       Next j
                    rsLockDialog.MoveNext
              Wend
        End If
        'second part ends here
        
        flxRetention.col = selcol
        flxRetention.row = selRow
        
        
        
        
   adoPay.Close
   adoconn.Close
   flxRetention.row = selRow
   flxRetention.col = selcol
   Set adoPay = Nothing
   Set adoconn = Nothing
End Function
Private Function IsPossible2Edit() As Boolean
  Dim adoPay As New ADODB.Recordset
   Dim adoRst As New ADODB.Recordset
   Dim adoconn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer
   Dim strTemp As String

   adoconn.Open getConnectionString
   szSQL = " SELECT BP.TRANS,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment BP WHERE MY_ID= '" & flxRetention.TextMatrix(flxRetention.row, 10) & "'"
   adoPay.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 Then  'szSQL <> UserSessionID shall be always true bcoz PI is generating only one seeion through out this module.
                flxRetention.col = 0
                'flxRetention.row = iPIEdit
                flxRetention.CellBackColor = vbRed
                strTemp = IIf(adoPay("TRANS").Value = "BP", "Bank Payment", "Bank Receipt")
                MsgBox "The selected " & strTemp & " is currently locked by '" & IIf(IsNull(adoPay("WindowsUserName").Value), "", adoPay("WindowsUserName").Value) & _
                "' on '" & IIf(IsNull(adoPay("MachineName").Value), "", adoPay("MachineName").Value) & "' in the '" & IIf(IsNull(adoPay("Module").Value), "", adoPay("Module").Value) & "'" & vbCrLf & "" & _
                        "screen for the Client '" & IIf(IsNull(adoPay("ClientID").Value), "", adoPay("ClientID").Value) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
                IsPossible2Edit = False
            Else 'lock row for this user in database
                flxRetention.col = 0
                IsPossible2Edit = True
                flxRetention.CellBackColor = vbWhite
            End If
   End If
   adoPay.Close
   Set adoPay = Nothing
   
   If IsPossible2Edit = True Then 'the reason I need to lock it here because it has passed all the tests here to open PI and now I can lock
             adoconn.Execute "Update tlbBankPayment Set  DateTimeStamp='" & Now & "',Module='Bank Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
                "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where MY_ID = '" & flxRetention.TextMatrix(flxRetention.row, 10) & "'"
   End If
   adoconn.Close
   Set adoconn = Nothing
End Function
Private Sub cmdEditBk_Click()
   Dim iCol As Integer
   Dim i As Integer
   Dim rCount As Integer
   Dim iIncDec As Integer
   Dim strStatementID As String
   For rCount = 1 To flxRetention.Rows - 1
        If flxRetention.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            iBankPayRow = rCount
        End If
   Next
   If iIncDec <> 1 Then 'bebugger shall not go further if you do not select one row
      MsgBox "Please select one transaction only.", vbInformation + vbOKOnly, "Transaction Selection"
      chkSelAll.Value = 0
      For i = 1 To flxRetention.Rows - 1
           If flxRetention.TextMatrix(i, 0) = "X" Then
                flxRetention.TextMatrix(i, 0) = ""
           End If
      Next i
      Exit Sub
   End If
   flxRetention.row = iBankPayRow
  
   strStatementID = flxRetention.TextMatrix(iBankPayRow, 6)
   If strStatementID <> "" Then
        MsgBox "It is not possible to edit a Retention if it has been applied to a Client Statement. ", vbInformation, "Warning"
        Exit Sub
   End If
   
   If strStatementID <> "" Then
        MsgBox " This Retention is assigned to a current Client Statement. You will need to run 'Modify Statement' on your current Client Statement for any changes made to be reflected in your Client Statement. ", vbInformation, "Warning"
   End If
   
   
   Load frmRetentionAdd
   frmRetentionAdd.Caption = "Edit Retention"
   
  bEditMode = True

   With frmRetentionAdd
        .txtClientList.Tag = flxRetention.TextMatrix(flxRetention.row, 2)
        .txtProperty.Tag = flxRetention.TextMatrix(flxRetention.row, 3)
        .txtProperty.text = findPropertyName(flxRetention.TextMatrix(flxRetention.row, 3))
        .txtBankCode.text = flxRetention.TextMatrix(flxRetention.row, 4) 'BankCode Bank Name
        .txtBankName = findBankName(flxRetention.TextMatrix(flxRetention.row, 2), flxRetention.TextMatrix(flxRetention.row, 4))
        'you need to amend bankcode here anol 2023-05-24
        .txtDate.text = flxRetention.TextMatrix(flxRetention.row, 8)
        .txtReference.text = flxRetention.TextMatrix(flxRetention.row, 9)
        .txtDetails.text = flxRetention.TextMatrix(flxRetention.row, 10)
        .txtNet.text = Format(flxRetention.TextMatrix(flxRetention.row, 11), "0.00")
        .szTransID = flxRetention.TextMatrix(flxRetention.row, 12)
        .txtClientList.text = flxRetention.TextMatrix(flxRetention.row, 13)
        .txtFund.Tag = flxRetention.TextMatrix(flxRetention.row, 14)
        .txtFund.text = flxRetention.TextMatrix(flxRetention.row, 15)
   End With
    LoadForm frmRetentionAdd
    frmRetentionAdd.txtBankBalance.Enabled = False
    frmRetentionAdd.txtRetention.Enabled = False
    frmRetentionAdd.txtAvailableBankBal.Enabled = False
    Call frmRetentionAdd.updateBankBalance
End Sub
Private Function findPropertyName(ID As String)
     Dim adoRst As New ADODB.Recordset
     Dim adoconn As New ADODB.Connection
     If ID = "" Then Exit Function
     Dim szSQL As String
     adoconn.Open getConnectionString
     adoRst.Open "Select PropertyName from Property where propertyID='" & ID & "'", adoconn, adOpenStatic, adLockReadOnly
     If Not adoRst.EOF Then
            findPropertyName = adoRst("PropertyName").Value
     End If
     adoRst.Close
     adoconn.Close
End Function
Private Function findBankName(ClientID As String, Code As String)
     Dim adoRst As New ADODB.Recordset
     Dim adoconn As New ADODB.Connection
     If Code = "" Then Exit Function
     Dim szSQL As String
     adoconn.Open getConnectionString
     adoRst.Open "Select Bank_AC_Name from tlbCLIENTBanks where CLIENT_ID='" & ClientID & "'  and NominalCode='" & Code & "'", adoconn, adOpenStatic, adLockReadOnly
     If Not adoRst.EOF Then
            findBankName = adoRst("Bank_AC_Name").Value
     End If
     adoRst.Close
     adoconn.Close
End Function
Public Sub LoadflxRetentionHist(adoconn As ADODB.Connection, Filter As String)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRst As New ADODB.Recordset
   Dim tempstr As String
   Call ConfigflxRetentionHist
'  Column Heading: "Empty|<SL No|<Date|<Bank|<PropName|<NC|<Fund|<Ref|<Desc|>Amount|TranID|PD"
    szSQL = "SELECT R.*,C.ClientName,F.FundName from RetentionDetails R,Client C,Fund F where " & _
            "R.isDeleted=false and ishist=true and R.ClientID=C.ClientID AND F.FundID=R.FundID Order by statementID,SLNumber "

   If Filter = "3" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
             szSQL = ""
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'
'    If Filter = "1" Then
'        If txtSearchNo.text <> "" Then
'            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
'            adoRst.Filter = "INvNo Like '%" & tempstr & "%'"
'        End If
'    End If
'
'
'
'
'    If Filter = "2" Then
'         If txtSearchRef.text <> "" Then
'            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
'            adoRst.Filter = "Details Like '%" & tempstr & "%'"
'        End If
'    End If
'   i = 1
   
    
    If adoRst.RecordCount = 0 Then
        flxRetentionHist.Clear
        flxRetentionHist.Rows = 2
    Else
        flxRetentionHist.Rows = adoRst.RecordCount + 1
    End If
    
    i = 1
   
   While Not adoRst.EOF
        flxRetentionHist.TextMatrix(i, 1) = "RT" & adoRst.Fields.Item("ID").Value
        flxRetentionHist.TextMatrix(i, 2) = adoRst.Fields.Item("ClientID").Value                           'Date
        flxRetentionHist.TextMatrix(i, 3) = IIf(IsNull(adoRst.Fields.Item("PropertyID").Value), "", adoRst.Fields.Item("PropertyID").Value)                           'Account Name
        flxRetentionHist.TextMatrix(i, 4) = IIf(IsNull(adoRst.Fields.Item("StatementID").Value), "", adoRst.Fields.Item("StatementID").Value) ' adoRst.Fields.Item("StatementID").Value   'Property Name
        flxRetentionHist.TextMatrix(i, 5) = IIf(IsNull(adoRst.Fields.Item("SlNumber").Value), "", adoRst.Fields.Item("SlNumber").Value) 'SL No
        flxRetentionHist.TextMatrix(i, 6) = adoRst.Fields.Item("RDate").Value   'NOMINAL_CODE
        flxRetentionHist.TextMatrix(i, 7) = adoRst.Fields.Item("Reference").Value                         'Fund
        flxRetentionHist.TextMatrix(i, 8) = adoRst.Fields.Item("Description").Value                             'Ref
        flxRetentionHist.TextMatrix(i, 9) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
        flxRetentionHist.TextMatrix(i, 10) = adoRst.Fields.Item("ID").Value
        flxRetentionHist.TextMatrix(i, 11) = adoRst.Fields.Item("ClientName").Value
        flxRetentionHist.TextMatrix(i, 12) = adoRst.Fields.Item("FundID").Value
        flxRetentionHist.TextMatrix(i, 13) = adoRst.Fields.Item("FundName").Value
        flxRetentionHist.RowHeight(i) = 280
        adoRst.MoveNext
      'If Not adoRst.EOF Then flxRetentionHist.AddItem ""
      i = i + 1
   Wend
'   If Len(colTransactionIDOtherPIGrid) > 0 Then
'        colTransactionIDOtherPIGrid = Left(colTransactionIDOtherPIGrid, Len(colTransactionIDOtherPIGrid) - 1)
'   End If
   adoRst.Close
   Set adoRst = Nothing
End Sub
Public Sub LoadflxRetention(adoconn As ADODB.Connection, Filter As String)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRst As New ADODB.Recordset
   Dim tempstr As String
   colTransactionIDOtherPIGrid = ""
'  Column Heading: "Empty|<SL No|<Date|<Bank|<PropName|<NC|<Fund|<Ref|<Desc|>Amount|TranID|PD"
    szSQL = "SELECT R.*,C.ClientName,F.FundName from RetentionDetails R,Client C,Fund F where " & _
            "R.isDeleted=false and ishist=false and R.ClientID=C.ClientID AND F.FundID=R.FundID Order by ID DESC,SLNumber "

   If Filter = "3" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
             szSQL = ""
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'
'    If Filter = "1" Then
'        If txtSearchNo.text <> "" Then
'            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
'            adoRst.Filter = "INvNo Like '%" & tempstr & "%'"
'        End If
'    End If
'
'
'
'
'    If Filter = "2" Then
'         If txtSearchRef.text <> "" Then
'            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
'            adoRst.Filter = "Details Like '%" & tempstr & "%'"
'        End If
'    End If
'   i = 1
   
    
    If adoRst.RecordCount = 0 Then
        flxRetention.Clear
        flxRetention.Rows = 2
    Else
        flxRetention.Rows = adoRst.RecordCount + 1
    End If
    
    i = 1
   Dim dblTotal As Double
   While Not adoRst.EOF
        flxRetention.TextMatrix(i, 1) = "RT" & adoRst.Fields.Item("ID").Value
        flxRetention.TextMatrix(i, 2) = adoRst.Fields.Item("ClientID").Value                           'Date
        flxRetention.TextMatrix(i, 3) = IIf(IsNull(adoRst.Fields.Item("PropertyID").Value), "", adoRst.Fields.Item("PropertyID").Value)                           'Account Name
        flxRetention.TextMatrix(i, 4) = IIf(IsNull(adoRst.Fields.Item("BankCode").Value), "", adoRst.Fields.Item("BankCode").Value)
        flxRetention.TextMatrix(i, 5) = findBankName(flxRetention.TextMatrix(i, 2), flxRetention.TextMatrix(i, 4))
        flxRetention.TextMatrix(i, 6) = IIf(IsNull(adoRst.Fields.Item("StatementID").Value), "", adoRst.Fields.Item("StatementID").Value) ' adoRst.Fields.Item("StatementID").Value   'Property Name
        flxRetention.TextMatrix(i, 7) = IIf(IsNull(adoRst.Fields.Item("SlNumber").Value), "", adoRst.Fields.Item("SlNumber").Value) 'SL No
        flxRetention.TextMatrix(i, 8) = adoRst.Fields.Item("RDate").Value   'NOMINAL_CODE
        flxRetention.TextMatrix(i, 9) = adoRst.Fields.Item("Reference").Value                         'Fund
        flxRetention.TextMatrix(i, 10) = adoRst.Fields.Item("Description").Value                             'Ref
        flxRetention.TextMatrix(i, 11) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
        dblTotal = dblTotal + Format(adoRst.Fields.Item("Amount").Value, "0.00")
        flxRetention.TextMatrix(i, 12) = adoRst.Fields.Item("ID").Value
        flxRetention.TextMatrix(i, 13) = adoRst.Fields.Item("ClientName").Value
        flxRetention.TextMatrix(i, 14) = adoRst.Fields.Item("FundID").Value
        flxRetention.TextMatrix(i, 15) = adoRst.Fields.Item("FundName").Value
        flxRetention.RowHeight(i) = 280
        flxRetention.ColAlignment(11) = vbRightJustify
        adoRst.MoveNext
 
      i = i + 1
   Wend
   txtRctTotal.text = Format(dblTotal, "0.00")
   adoRst.Close
   Set adoRst = Nothing
End Sub

Public Sub LoadflxRetentionbyclient(adoconn As ADODB.Connection)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRst As New ADODB.Recordset

'  Column Heading: "Empty|<SL No|<Date|<Bank|<PropName|<NC|<Fund|<Ref|<Desc|>Amount|TranID|PD"
'                     ^      ^      ^     ^      ^      ^     ^    ^     ^       ^     ^
   szSQL = "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, BP.MY_ID, C.CLIENT_ID, " & _
                  "TT.DESCRIPTION AS Type1, BP.PROJ_REF AS Rfn, BP.PropertyID, BP.UNIT_ID, BP.NOMINAL_CODE, " & _
                  "BP.BANK_AC AS ACC, C.Bank_AC_Name, BP.DEPT_ID, BP.NET_AMOUNT, BP.VAT, BP.ReconNow, " & _
                  "BP.TAX_CODE, P.PropertyName, BP.PostingDate " & _
           "FROM ((tlbBankPayment AS BP INNER JOIN tlbTransactionTypes AS TT ON BP.TransactionType = TT.TYPE_ID) INNER JOIN " & _
                  "tlbClientBanks AS C ON BP.BANK_AC = C.NominalCode) LEFT JOIN " & _
                  "Property AS P ON BP.PropertyID = P.PropertyID " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND tlbClientBanks.ClientID='" & txtClientList.Tag & "' AND " & _
                  "C.CLIENT_ID = BP.ClientID AND " & _
                  "BP.UPDATE_SAGE = FALSE " & _
           "ORDER BY TRANS, CLNG(TRAN_ID)"

'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   i = 1
   flxRetention.Clear
   flxRetention.Rows = 2
   While Not adoRst.EOF
      flxRetention.TextMatrix(i, 1) = adoRst.Fields.Item("Type2").Value & _
                                    adoRst.Fields.Item("T_ID").Value                            'SL No
      flxRetention.TextMatrix(i, 2) = adoRst.Fields.Item("RDate").Value                           'Date
      flxRetention.TextMatrix(i, 3) = adoRst.Fields.Item("ACC").Value                             'Account Name
      flxRetention.TextMatrix(i, 4) = IIf(IsNull(adoRst!PropertyName), "", adoRst!PropertyName)   'Property Name
      flxRetention.TextMatrix(i, 5) = IIf(IsNull(adoRst!Nominal_code), "", adoRst!Nominal_code)   'NOMINAL_CODE
      flxRetention.TextMatrix(i, 6) = adoRst.Fields.Item("DEPT_ID").Value                         'Fund
      flxRetention.TextMatrix(i, 7) = adoRst.Fields.Item("Rfn").Value                             'Ref
      flxRetention.TextMatrix(i, 8) = adoRst.Fields.Item("Details").Value                         'Desc
      flxRetention.TextMatrix(i, 9) = Format(adoRst.Fields.Item("Amount").Value, "0.00")          'Amt
      flxRetention.TextMatrix(i, 10) = adoRst.Fields.Item("MY_ID").Value                          'TranID
      flxRetention.TextMatrix(i, 11) = IIf(IsNull(adoRst!propertyID), "", adoRst!propertyID)      'PropertyID
      flxRetention.TextMatrix(i, 12) = IIf(IsNull(adoRst!UNIT_ID), "", adoRst!UNIT_ID)            'UNIT_ID
      flxRetention.TextMatrix(i, 13) = IIf(IsNull(adoRst!CLIENT_ID), "", adoRst!CLIENT_ID)        'CLIENT ID
      flxRetention.TextMatrix(i, 14) = IIf(IsNull(adoRst!ACC), "", adoRst!ACC)                    'BANK_AC
      flxRetention.TextMatrix(i, 15) = adoRst.Fields.Item("NET_AMOUNT").Value                     'NET_AMOUNT
      flxRetention.TextMatrix(i, 16) = adoRst.Fields.Item("VAT").Value                            'VAT
      flxRetention.TextMatrix(i, 17) = IIf(IsNull(adoRst!ReconNow), "N", "Y")
      flxRetention.TextMatrix(i, 18) = IIf(IsNull(adoRst!TAX_CODE), "", adoRst!TAX_CODE)
      flxRetention.TextMatrix(i, 19) = IIf(IsNull(adoRst!postingDate), adoRst.Fields.Item("RDate").Value, adoRst!postingDate)    'Posting Date

      adoRst.MoveNext
      If Not adoRst.EOF Then flxRetention.AddItem ""
      i = i + 1
   Wend
   adoRst.Close
   Set adoRst = Nothing
End Sub
Public Sub LoadflxRetentionbyProperty(adoconn As ADODB.Connection)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRst As New ADODB.Recordset

'  Column Heading: "Empty|<SL No|<Date|<Bank|<PropName|<NC|<Fund|<Ref|<Desc|>Amount|TranID|PD"
'                     ^      ^      ^     ^      ^      ^     ^    ^     ^       ^     ^
   szSQL = "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, BP.MY_ID, C.CLIENT_ID, " & _
                  "TT.DESCRIPTION AS Type1, BP.PROJ_REF AS Rfn, BP.PropertyID, BP.UNIT_ID, BP.NOMINAL_CODE, " & _
                  "BP.BANK_AC AS ACC, C.Bank_AC_Name, BP.DEPT_ID, BP.NET_AMOUNT, BP.VAT, BP.ReconNow, " & _
                  "BP.TAX_CODE, P.PropertyName, BP.PostingDate " & _
           "FROM ((tlbBankPayment AS BP INNER JOIN tlbTransactionTypes AS TT ON BP.TransactionType = TT.TYPE_ID) INNER JOIN " & _
                  "tlbClientBanks AS C ON BP.BANK_AC = C.NominalCode) LEFT JOIN " & _
                  "Property AS P ON BP.PropertyID = P.PropertyID " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND tlbClientBanks.PropertyID='" & txtPoperty.Tag & "' AND " & _
                  "C.CLIENT_ID = BP.ClientID AND " & _
                  "BP.UPDATE_SAGE = FALSE " & _
           "ORDER BY TRANS, CLNG(TRAN_ID)"

'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   i = 1
   flxRetention.Clear
   flxRetention.Rows = 2
   While Not adoRst.EOF
      flxRetention.TextMatrix(i, 1) = adoRst.Fields.Item("Type2").Value & _
                                    adoRst.Fields.Item("T_ID").Value                            'SL No
      flxRetention.TextMatrix(i, 2) = adoRst.Fields.Item("RDate").Value                           'Date
      flxRetention.TextMatrix(i, 3) = adoRst.Fields.Item("ACC").Value                             'Account Name
      flxRetention.TextMatrix(i, 4) = IIf(IsNull(adoRst!PropertyName), "", adoRst!PropertyName)   'Property Name
      flxRetention.TextMatrix(i, 5) = IIf(IsNull(adoRst!Nominal_code), "", adoRst!Nominal_code)   'NOMINAL_CODE
      flxRetention.TextMatrix(i, 6) = adoRst.Fields.Item("DEPT_ID").Value                         'Fund
      flxRetention.TextMatrix(i, 7) = adoRst.Fields.Item("Rfn").Value                             'Ref
      flxRetention.TextMatrix(i, 8) = adoRst.Fields.Item("Details").Value                         'Desc
      flxRetention.TextMatrix(i, 9) = Format(adoRst.Fields.Item("Amount").Value, "0.00")          'Amt
      flxRetention.TextMatrix(i, 10) = adoRst.Fields.Item("MY_ID").Value                          'TranID
      flxRetention.TextMatrix(i, 11) = IIf(IsNull(adoRst!propertyID), "", adoRst!propertyID)      'PropertyID
      flxRetention.TextMatrix(i, 12) = IIf(IsNull(adoRst!UNIT_ID), "", adoRst!UNIT_ID)            'UNIT_ID
      flxRetention.TextMatrix(i, 13) = IIf(IsNull(adoRst!CLIENT_ID), "", adoRst!CLIENT_ID)        'CLIENT ID
      flxRetention.TextMatrix(i, 14) = IIf(IsNull(adoRst!ACC), "", adoRst!ACC)                    'BANK_AC
      flxRetention.TextMatrix(i, 15) = adoRst.Fields.Item("NET_AMOUNT").Value                     'NET_AMOUNT
      flxRetention.TextMatrix(i, 16) = adoRst.Fields.Item("VAT").Value                            'VAT
      flxRetention.TextMatrix(i, 17) = IIf(IsNull(adoRst!ReconNow), "N", "Y")
      flxRetention.TextMatrix(i, 18) = IIf(IsNull(adoRst!TAX_CODE), "", adoRst!TAX_CODE)
      flxRetention.TextMatrix(i, 19) = IIf(IsNull(adoRst!postingDate), adoRst.Fields.Item("RDate").Value, adoRst!postingDate)    'Posting Date

      adoRst.MoveNext
      If Not adoRst.EOF Then flxRetention.AddItem ""
      i = i + 1
   Wend
   adoRst.Close
   Set adoRst = Nothing
End Sub
'Private Sub LoadBankAccountInCombo(ByVal adoconn As ADODB.Connection)
'   On Error GoTo Error_Handler
'
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, Data() As String, j As Integer
'   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer
'
'   szSQL = "SELECT C.NominalCode AS BNC, N.Name AS BNN " & _
'           "FROM tlbClientBanks AS C, NominalLedger AS N " & _
'           "WHERE C.CLIENT_ID = N.ClientID AND C.NominalCode = N.Code AND C.CLIENT_ID <> '';"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   iTotalRow = adoRst.RecordCount
'   iTotalCol = adoRst.Fields.Count
'   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String
'
'   For i = 0 To iTotalRow
'       For j = 0 To iTotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboBC.Column() = Data()
'
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

Private Sub fraButtons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
   
   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub lblBankRec_Click(Index As Integer)
    If Index >= 0 And Index <= 6 Then
'         Me.MousePointer = vbHourglass
'       If Index = 0 Then
'             SortingGrid flxRetention, Index + 1, bSortingCol(Index), "Integer"
'      Else
'            SortingGrid flxRetention, Index + 1, bSortingCol(Index)
'      End If
'
'      bSortingCol(Index) = IIf(bSortingCol(Index), False, True)
'
       ' LblSortingClicked Index, lblBankRec, 0, 6
'        Me.MousePointer = vbArrow
            lblBankRec(Index).FontBold = Not lblBankRec(Index).FontBold
            Dim adoconn As New ADODB.Connection
            adoconn.Open getConnectionString
            ConfigflxRetention
            LoadflxRetention1 adoconn, Index, IIf(lblBankRec(Index).FontBold, "DESC", "ASC")
            adoconn.Close
    End If
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

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDisplay.SetFocus
    End If
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub
Private Sub txtDateFrom_LostFocus()
   Dim X As Boolean

   X = TextBoxFormatDate(txtDateFrom)

   If X And txtDateTo.text = "" Then txtDateTo.text = txtDateFrom.text
End Sub
Public Sub LoadflxRetention1(adoconn As ADODB.Connection, j As Integer, strSort As String)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRst As New ADODB.Recordset
   Dim strWhereBC As String
   Dim strWhereType As String
    Dim szOrderby, strWhereCL, strWherePR, strWhereDt As String
    If j = 0 Then
        szOrderby = "ORDER BY statementID, SLNumber " & strSort
    End If
    If j = 1 Then
        szOrderby = "ORDER BY statementID,R.RDATE " & strSort
    End If
    If j = 2 Then
       ' szOrderby = "ORDER BY statementID,R.BANK_AC " & strSort
    End If
    If j = 3 Then
        szOrderby = "ORDER BY TRANS,R.PropertyID " & strSort
    End If
    If j = 4 Then
       ' szOrderby = "ORDER BY statementID,BP.NOMINAL_CODE " & strSort
    End If
    If j = 5 Then
       ' szOrderby = "ORDER BY statementID,BP.DEPT_ID " & strSort
    End If
    If j = 6 Then
       ' szOrderby = "ORDER BY statementID,BP.PROJ_REF " & strSort
    End If
    
    If txtClientList.Tag <> "ALL" Then
        strWhereCL = " AND R.ClientID='" & txtClientList.Tag & "' "
    End If

    If txtPoperty.Tag <> "ALL" Then
        strWherePR = " AND R.PropertyID='" & txtPoperty.Tag & "' "
    End If
    If Trim(txtDateFrom.text) <> "" And txtDateTo.text <> "" Then

        strWhereDt = " AND R.RDATE>=#" & Format(txtDateFrom.text, "dd mmmm yyyy") & "# AND R.RDATE<=#" & Format(txtDateTo.text, "dd mmmm yyyy") & "# "
    End If
   szSQL = "SELECT R.*,C.ClientName,F.FundName from RetentionDetails R,Client C,Fund F where " & _
            "R.isDeleted=false and ishist=false and R.ClientID=C.ClientID AND F.FundID=R.FundID " & strWhereCL & strWherePR & " Order by statementID,SLNumber "


'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   i = 1
   flxRetention.Clear
   flxRetention.Rows = 2
   Dim dblTotal As Double
   ConfigflxRetention
    While Not adoRst.EOF
        flxRetention.TextMatrix(i, 1) = "RT" & adoRst.Fields.Item("ID").Value
        flxRetention.TextMatrix(i, 2) = adoRst.Fields.Item("ClientID").Value                           'Date
        flxRetention.TextMatrix(i, 3) = IIf(IsNull(adoRst.Fields.Item("PropertyID").Value), "", adoRst.Fields.Item("PropertyID").Value)                           'Account Name
        flxRetention.TextMatrix(i, 4) = IIf(IsNull(adoRst.Fields.Item("BankCode").Value), "", adoRst.Fields.Item("BankCode").Value)                           'Account Name
        flxRetention.TextMatrix(i, 5) = IIf(IsNull(adoRst.Fields.Item("BankCode").Value), "", adoRst.Fields.Item("BankCode").Value)
        flxRetention.TextMatrix(i, 6) = IIf(IsNull(adoRst.Fields.Item("StatementID").Value), "", adoRst.Fields.Item("StatementID").Value) ' adoRst.Fields.Item("StatementID").Value   'Property Name
        flxRetention.TextMatrix(i, 7) = IIf(IsNull(adoRst.Fields.Item("SlNumber").Value), "", adoRst.Fields.Item("SlNumber").Value) 'SL No
        flxRetention.TextMatrix(i, 8) = adoRst.Fields.Item("RDate").Value   'NOMINAL_CODE
        flxRetention.TextMatrix(i, 9) = adoRst.Fields.Item("Reference").Value                         'Fund
        flxRetention.TextMatrix(i, 10) = adoRst.Fields.Item("Description").Value                             'Ref
        flxRetention.TextMatrix(i, 11) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
        dblTotal = dblTotal + Format(adoRst.Fields.Item("Amount").Value, "0.00")
        flxRetention.TextMatrix(i, 12) = adoRst.Fields.Item("ID").Value
        flxRetention.TextMatrix(i, 13) = adoRst.Fields.Item("ClientName").Value
        flxRetention.TextMatrix(i, 14) = adoRst.Fields.Item("FundID").Value
        flxRetention.TextMatrix(i, 15) = adoRst.Fields.Item("FundName").Value
        flxRetention.RowHeight(i) = 280
        
        adoRst.MoveNext
      If Not adoRst.EOF Then flxRetention.AddItem ""
      i = i + 1
   Wend
   txtRctTotal.text = Format(dblTotal, "0.00")
   adoRst.Close
   Set adoRst = Nothing
End Sub
Private Sub tabPayment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
   
   If fraAddTrans.Visible Then fraAddTrans.Visible = False
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
Private Sub txtSearchClientID_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
              flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
              flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           flxClient.SetFocus
    End If
    If KeyCode = 13 Then
      txtSearchClientName.SetFocus
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
         picClient.Visible = False
          tabPayment.Enabled = True
          
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           End If
    End If
End Sub

Private Sub txtSearchClientName_Change()
   'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
            flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
         flxClient.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            flxClient.SetFocus
        End If
    End If
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 27 Then
         picClient.Visible = False
          tabPayment.Enabled = True
          
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           End If
    End If
End Sub

Private Sub txtSearchNo_Change()
    txtSearchFromD.text = ""
    txtSearchToD.text = ""
    txtSearchRef.text = ""
    Dim adoconn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        adoconn.Open getConnectionString
        If Len(txtSearchNo.text) > 0 Then
            LoadflxRetention adoconn, "1"
        Else
            LoadflxRetention adoconn, ""
        End If
    '    fmeLoading.Visible = False
        adoconn.Close
        Set adoconn = Nothing
        If Len(txtSearchNo.text) > 0 Then
            cmdSearch.Caption = "Clear Sea&rch"
        Else
            cmdSearch.Caption = "Sea&rch"
        End If
    ElseIf tabPayment.Tab = 1 Then
        adoconn.Open getConnectionString
        If Len(txtSearchNo.text) > 0 Then
            LoadflxRetentionHist adoconn, "1"
        Else
            LoadflxRetentionHist adoconn, ""
        End If
        adoconn.Close
        Set adoconn = Nothing
        If Len(txtSearchNo.text) > 0 Then
            cmdSearchHistory.Caption = "Clear Sea&rch"
        Else
            cmdSearchHistory.Caption = "Sea&rch"
        End If
     End If
End Sub

Private Sub txtSearchRef_Change()
     txtSearchFromD.text = ""
    txtSearchToD.text = ""
    txtSearchNo.text = ""
    Dim adoconn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        adoconn.Open getConnectionString
        If Len(txtSearchRef.text) > 0 Then
            LoadflxRetention adoconn, "2"
        Else
            LoadflxRetention adoconn, ""
        End If
    '    fmeLoading.Visible = False
        adoconn.Close
        Set adoconn = Nothing
        If Len(txtSearchRef.text) > 0 Then
            cmdSearch.Caption = "Clear Sea&rch"
        Else
            cmdSearch.Caption = "Sea&rch"
        End If
    ElseIf tabPayment.Tab = 1 Then
        adoconn.Open getConnectionString
        If Len(txtSearchRef.text) > 0 Then
            LoadflxRetentionHist adoconn, "2"
        Else
            LoadflxRetentionHist adoconn, ""
        End If
        adoconn.Close
        Set adoconn = Nothing
        If Len(txtSearchRef.text) > 0 Then
            cmdSearchHistory.Caption = "Clear Sea&rch"
        Else
            cmdSearchHistory.Caption = "Sea&rch"
        End If
     End If
End Sub
