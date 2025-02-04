VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBankTransactions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Receipt and Payment"
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16575
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankTransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   16575
   Begin VB.PictureBox picBankCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2070
      ScaleHeight     =   3180
      ScaleWidth      =   4035
      TabIndex        =   87
      Top             =   1710
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
         TabIndex        =   88
         Top             =   90
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBankCode 
         Height          =   2715
         Left            =   45
         TabIndex        =   89
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
         TabIndex        =   91
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
         TabIndex        =   90
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
      TabIndex        =   70
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
         TabIndex        =   71
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
            TabIndex        =   76
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtSearchToD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2025
            MaxLength       =   80
            TabIndex        =   75
            Top             =   1125
            Width           =   1380
         End
         Begin VB.TextBox txtSearchFromD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   80
            TabIndex        =   74
            Top             =   1125
            Width           =   1290
         End
         Begin VB.TextBox txtSearchRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   73
            Top             =   790
            Width           =   2685
         End
         Begin VB.TextBox txtSearchNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   10
            TabIndex        =   72
            Top             =   450
            Width           =   2685
         End
         Begin VB.CommandButton cmdSearchCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   2055
            TabIndex        =   79
            Top             =   1635
            Width           =   1200
         End
         Begin VB.CommandButton cmdSearchOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   77
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   78
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
      TabIndex        =   55
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
         TabIndex        =   56
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   57
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   58
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
      TabIndex        =   46
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
         TabIndex        =   47
         Top             =   40
         Width           =   3615
         Begin VB.CommandButton cmdCancelTrans 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   2100
            TabIndex        =   52
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddTrans 
            Caption         =   "&OK"
            Height          =   315
            Left            =   80
            TabIndex        =   51
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optBankReceipt 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Bank Receipt"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   735
            Width           =   1680
         End
         Begin VB.OptionButton optBankPayment 
            BackColor       =   &H00E5E5E5&
            Caption         =   "Bank Payment"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   1335
            TabIndex        =   48
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
            TabIndex        =   53
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
      TabIndex        =   15
      Top             =   0
      Width           =   16515
      _ExtentX        =   29131
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
      TabCaption(0)   =   "Bank &Receipt && Payment"
      TabPicture(0)   =   "frmBankTransactions.frx":000C
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
      Tab(0).Control(5)=   "Label3(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboBC"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblBankRec(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblBankRec(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblBankRec(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblBankRec(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblBankRec(6)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblBankRec(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblBankRec(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblBankRec(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblBankRec(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtClientList"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPoperty"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkSelAll"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label3(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label3(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "fraButtons"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "flxBankPay"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdClientList"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdProperty"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtDateFrom"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtDateTo"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdDisplay"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtTransTypeFilter"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdTransTypeFilter"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtBankAccountFilter"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdBankAccountFiilter"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Bank Receipt && Payment &History"
      TabPicture(1)   =   "frmBankTransactions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSearchHistory"
      Tab(1).Control(1)=   "cmdReverseHistory"
      Tab(1).Control(2)=   "cmdEdit"
      Tab(1).Control(3)=   "cmdClosebk(1)"
      Tab(1).Control(4)=   "flxBankPayHist"
      Tab(1).Control(5)=   "chkallHist"
      Tab(1).Control(6)=   "lblBankHist(4)"
      Tab(1).Control(7)=   "lblBankHist(10)"
      Tab(1).Control(8)=   "lblBankHist(9)"
      Tab(1).Control(9)=   "lblBankHist(8)"
      Tab(1).Control(10)=   "lblBankHist(6)"
      Tab(1).Control(11)=   "lblBankHist(7)"
      Tab(1).Control(12)=   "lblBankHist(5)"
      Tab(1).Control(13)=   "lblBankHist(3)"
      Tab(1).Control(14)=   "lblBankHist(2)"
      Tab(1).Control(15)=   "lblBankHist(1)"
      Tab(1).Control(16)=   "lblBankHist(0)"
      Tab(1).Control(17)=   "Shape4(0)"
      Tab(1).ControlCount=   18
      Begin VB.CommandButton cmdBankAccountFiilter 
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
         Left            =   11520
         TabIndex        =   11
         Top             =   540
         Width           =   300
      End
      Begin VB.TextBox txtBankAccountFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10755
         TabIndex        =   10
         Text            =   "ALL"
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdTransTypeFilter 
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
         Left            =   9270
         TabIndex        =   9
         Top             =   540
         Width           =   300
      End
      Begin VB.TextBox txtTransTypeFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8055
         TabIndex        =   8
         Text            =   "ALL"
         Top             =   540
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearchHistory 
         Caption         =   "Sea&rch"
         Height          =   400
         Left            =   -72705
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   7110
         Width           =   1395
      End
      Begin VB.CommandButton cmdReverseHistory 
         Caption         =   "&Reverse History"
         Height          =   400
         Left            =   -74730
         TabIndex        =   67
         Top             =   7110
         Width           =   1935
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Display"
         Height          =   430
         Left            =   15480
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   450
         Width           =   960
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   14355
         TabIndex        =   13
         Top             =   540
         Width           =   1065
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12645
         TabIndex        =   12
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
         Left            =   6840
         TabIndex        =   7
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
         Left            =   3555
         TabIndex        =   6
         Top             =   540
         Width           =   300
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Transaction"
         Height          =   400
         Left            =   -70770
         TabIndex        =   33
         Top             =   7110
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdClosebk 
         Caption         =   "&Close"
         Height          =   400
         Index           =   1
         Left            =   -63060
         TabIndex        =   32
         Top             =   7080
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankPay 
         Height          =   7890
         Left            =   2040
         TabIndex        =   16
         Top             =   1365
         Width           =   14430
         _ExtentX        =   25453
         _ExtentY        =   13917
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
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
            TabIndex        =   69
            Top             =   5085
            Width           =   1395
         End
         Begin VB.CommandButton cmdPostHistory 
            Caption         =   "Post to &History"
            Height          =   430
            Left            =   120
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   3000
            Width           =   1450
         End
         Begin VB.CommandButton cmdPrintList 
            Caption         =   "&Print List"
            Height          =   430
            Left            =   120
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2400
            Width           =   1450
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "&Copy"
            Height          =   430
            Left            =   135
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   1575
            Visible         =   0   'False
            Width           =   1450
         End
         Begin VB.CommandButton cmdEditBk 
            Caption         =   "&Edit"
            Height          =   430
            Left            =   120
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   900
            Width           =   1450
         End
         Begin VB.CommandButton cmdNewBk_ 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add New"
            Height          =   430
            Left            =   1560
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   210
            Visible         =   0   'False
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
            TabIndex        =   5
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
            Height          =   1335
            Left            =   0
            Top             =   2280
            Width           =   1695
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankPayHist 
         Height          =   6315
         Left            =   -74880
         TabIndex        =   34
         Top             =   705
         Width           =   13155
         _ExtentX        =   23204
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account :"
         Height          =   195
         Index           =   1
         Left            =   9645
         TabIndex        =   86
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Type :"
         Height          =   195
         Index           =   0
         Left            =   7215
         TabIndex        =   85
         Top             =   585
         Width           =   825
      End
      Begin MSForms.CheckBox chkallHist 
         Height          =   255
         Left            =   -74865
         TabIndex        =   84
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
         TabIndex        =   83
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
         Left            =   13725
         TabIndex        =   66
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
         Left            =   11835
         TabIndex        =   65
         Top             =   585
         Width           =   765
      End
      Begin MSForms.TextBox txtPoperty 
         Height          =   285
         Left            =   4635
         TabIndex        =   64
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
         Left            =   1035
         TabIndex        =   54
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
         Caption         =   "NC"
         Height          =   255
         Index           =   4
         Left            =   -70320
         TabIndex        =   45
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   255
         Index           =   10
         Left            =   -63090
         TabIndex        =   44
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
         Height          =   255
         Index           =   9
         Left            =   -63810
         TabIndex        =   43
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Net"
         Height          =   255
         Index           =   8
         Left            =   -64800
         TabIndex        =   42
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   255
         Index           =   6
         Left            =   -68280
         TabIndex        =   41
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
         Height          =   255
         Index           =   7
         Left            =   -66840
         TabIndex        =   40
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   255
         Index           =   5
         Left            =   -69600
         TabIndex        =   39
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   255
         Index           =   3
         Left            =   -72060
         TabIndex        =   38
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Index           =   2
         Left            =   -73140
         TabIndex        =   37
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         Height          =   255
         Index           =   1
         Left            =   -73860
         TabIndex        =   36
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblBankHist 
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   35
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   255
         Index           =   8
         Left            =   14220
         TabIndex        =   31
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   2280
         TabIndex        =   30
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   4170
         TabIndex        =   29
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Property Name"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   5130
         TabIndex        =   28
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   8655
         TabIndex        =   27
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Index           =   7
         Left            =   9975
         TabIndex        =   26
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   6765
         TabIndex        =   25
         Top             =   1125
         Width           =   615
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   7815
         TabIndex        =   24
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label lblBankRec 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   23
         Top             =   1125
         Width           =   375
      End
      Begin MSForms.ComboBox cboBC 
         Height          =   285
         Left            =   15300
         TabIndex        =   22
         Top             =   675
         Visible         =   0   'False
         Width           =   2745
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4842;503"
         TextColumn      =   2
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
         Object.Width           =   "1058;3527"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank A/C:"
         Height          =   195
         Index           =   7
         Left            =   11070
         TabIndex        =   21
         Top             =   990
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   660
         Index           =   3
         Left            =   240
         Top             =   360
         Width           =   16230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   6
         Left            =   3945
         TabIndex        =   20
         Top             =   555
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   5
         Left            =   450
         TabIndex        =   19
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
         Width           =   16230
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   -74880
         Top             =   480
         Width           =   13155
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
         Width           =   14430
      End
   End
End
Attribute VB_Name = "frmBankTransactions"
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
      For iRow = 1 To flxBankPayHist.Rows - 1
         flxBankPayHist.TextMatrix(iRow, 0) = ""
      Next iRow
   Else
      For iRow = 1 To flxBankPayHist.Rows - 1
         flxBankPayHist.TextMatrix(iRow, 0) = "X"
      Next iRow
   End If
End Sub

Private Sub chkSelAll_Click()
    Dim iRow As Integer
   
   If Not chkSelAll.Value Then
      For iRow = 1 To flxBankPay.Rows - 1
         flxBankPay.TextMatrix(iRow, 0) = ""
      Next iRow
   Else
      For iRow = 1 To flxBankPay.Rows - 1
         flxBankPay.TextMatrix(iRow, 0) = "X"
      Next iRow
   End If
End Sub

Private Sub cmdBankAccountFiilter_Click()
    sText = "1"
    picBankCode.Visible = True
    picBankCode.Top = cmdBankAccountFiilter.Top
    picBankCode.Left = Label3(1).Left
    Call LoadBankAccount
End Sub

Private Sub cmdBankClose_Click()
    picBankCode.Visible = False
End Sub

Private Sub cmdSearchCancel_Click()
    Dim adoConn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        fraSearch.Visible = False
        adoConn.Open getConnectionString
        If cmdSearch.Caption = "Clear Sea&rch" Then
             cmdSearch.Caption = "Sea&rch"
        End If
        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
        Else
             Call LoadFlxBankPay(adoConn, "")
        End If
        adoConn.Close
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
    Dim adoConn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        adoConn.Open getConnectionString
        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
            LoadFlxBankPay adoConn, ""
            'fmeLoading.Visible = False      cmdSearch.Caption = "Sea&rch"
        ElseIf Trim(txtSearchNo.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchRef.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
            LoadFlxBankPay adoConn, "3"
            cmdSearch.Caption = "Clear Sea&rch"
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
            cmdSearch.Caption = "Clear Sea&rch"
            If tabPayment.Tab = 0 Then
                LoadFlxBankPay adoConn, "3"
            Else
                LoadflxBankPayHist adoConn, "3"
            End If
        End If
        adoConn.Close
        Set adoConn = Nothing
    End If
End Sub

Private Sub cmdTransTypeFilter_Click()
    sText = "2"
    picBankCode.Visible = True
    picBankCode.Top = cmdTransTypeFilter.Top
    picBankCode.Left = Label3(0).Left
    Call LoadBankTranType
End Sub
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
Private Sub gridBankCode_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If sText = "1" Then
        txtBankAccountFilter.text = gridBankCode.TextMatrix(gridBankCode.row, 1)
        picBankCode.Visible = False
        LoadFlxBankPay1 adoConn, 0, "ASC"
    ElseIf sText = "2" Then
        txtTransTypeFilter.Tag = gridBankCode.TextMatrix(gridBankCode.row, 1)
        txtTransTypeFilter.text = gridBankCode.TextMatrix(gridBankCode.row, 2)
        picBankCode.Visible = False
        LoadFlxBankPay1 adoConn, 0, "ASC"
    End If
    adoConn.Close
    Set adoConn = Nothing
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
'        LoadFlxBankPaybyclient adoconn
'   Else
'        LoadFlxBankPay adoconn
'   End If
'   loadProperty adoconn
'   adoconn.Close
End Sub

Private Sub SortTheGrid()
  ' If cmbProperty.ListCount <= 0 Then Exit Sub

   Dim i As Integer

   If txtClientList.Tag <> "ALL" Then
      For i = 1 To flxBankPay.Rows - 1
         If flxBankPay.TextMatrix(i, 13) = txtClientList.Tag Then
            flxBankPay.RowHeight(i) = 240
         Else
            flxBankPay.RowHeight(i) = 0
         End If
      Next i
   Else
      For i = 1 To flxBankPay.Rows - 1
         flxBankPay.RowHeight(i) = 240
      Next i
   End If

   If txtPoperty.Tag <> "ALL" Then
      For i = 1 To flxBankPay.Rows - 1
         If flxBankPay.TextMatrix(i, 11) = txtPoperty.Tag And flxBankPay.RowHeight(i) = 240 Then
            flxBankPay.RowHeight(i) = 240
         Else
            flxBankPay.RowHeight(i) = 0
         End If
      Next i
   End If

   If cboBC.Value <> "" Then
      For i = 1 To flxBankPay.Rows - 1
         If flxBankPay.TextMatrix(i, 14) = cboBC.Value And flxBankPay.RowHeight(i) = 240 Then
            flxBankPay.RowHeight(i) = 240
         Else
            flxBankPay.RowHeight(i) = 0
         End If
      Next i
   End If
End Sub

Private Sub cmbProperty_Click()
   SortTheGrid
'    Dim adoconn As New ADODB.Connection
'   adoconn.Open getConnectionString
'   If cmbProperty.text <> "ALL" Then
'        LoadFlxBankPaybyProperty adoconn
'   Else
'        LoadFlxBankPaybyclient adoconn
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

   Dim adoConn As New ADODB.Connection
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

   
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub cmdCloseBk_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdCloseSearch_Click()
    fraSearch.Visible = False
End Sub

Private Sub cmdCopy_Click()
'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + fraButtons.Top + cmdCopy.Top + frmBankTransactions.Top + tabPayment.Top + 1150
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + fraButtons.Left + frmBankTransactions.Left + tabPayment.Left + cmdCopy.Left + 80
   frmPopUpMenu.CallingFrom "BankCopy"
   frmPopUpMenu.Show
End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub cmdDisplay_Click()
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        
        
        If cmdDisplay.Caption = "&Display" And txtDateFrom.text <> "" And txtDateTo.text <> "" Then
             cmdDisplay.Caption = "&Clear"
        Else
            cmdDisplay.Caption = "&Display"
            txtDateFrom.text = ""
            txtDateTo.text = ""
        End If
        ConfigFlxBankPay
        LoadFlxBankPay1 adoConn, 0, "ASC"
        adoConn.Close
End Sub

Private Sub cmdEdit_Click()
   If flxBankPayHist.TextMatrix(flxBankPayHist.row, 0) = "" Then
      Exit Sub
   End If
   If flxBankPayHist.TextMatrix(flxBankPayHist.row, 15) = "Y" Then
      MsgBox "The transaction has been bank reconciled."
      Exit Sub
   End If
   Load frmBankTranEdit
   frmBankTranEdit.Caption = "Edit Bank Transaction"
   frmBankTranEdit.FrmBankTranEdit_CALLING_FROM = Me.Name & "History"

   With frmBankTranEdit
      .szTransID = flxBankPayHist.TextMatrix(flxBankPayHist.row, 11)
      .txtClientList.Tag = flxBankPayHist.TextMatrix(flxBankPayHist.row, 12)
      .txtClientList.text = flxBankPayHist.TextMatrix(flxBankPayHist.row, 18)
      .txtBankCode.text = flxBankPayHist.TextMatrix(flxBankPayHist.row, 1)
      .txtBankName.text = flxBankPayHist.TextMatrix(flxBankPayHist.row, 19)
      .txtProperty.Tag = flxBankPayHist.TextMatrix(flxBankPayHist.row, 13)
      .txtUnit.Tag = flxBankPayHist.TextMatrix(flxBankPayHist.row, 16)
     
      .txtDetails.text = flxBankPayHist.TextMatrix(flxBankPayHist.row, 7)
      .txtReference.text = flxBankPayHist.TextMatrix(flxBankPayHist.row, 6)
      .txtNC.Tag = flxBankPayHist.TextMatrix(flxBankPayHist.row, 4)
      .txtNC.text = flxBankPayHist.TextMatrix(flxBankPayHist.row, 20)
      .txtFund.Tag = flxBankPayHist.TextMatrix(flxBankPayHist.row, 14)
      .txtFund.Tag = flxBankPayHist.TextMatrix(flxBankPayHist.row, 6)
      .txtDate.text = flxBankPayHist.TextMatrix(flxBankPayHist.row, 2)
      .txtNet.text = Format(flxBankPayHist.TextMatrix(flxBankPayHist.row, 8), "0.00")
      '.cboVat.Value = flxBankPayHist.TextMatrix(flxBankPayHist.row, 17)
      .Label1(24).Caption = flxBankPayHist.TextMatrix(flxBankPayHist.row, 17)
      .txtVat_.text = Format(Val(flxBankPayHist.TextMatrix(flxBankPayHist.row, 9)), "0.00")
      .txtTotal.text = Format(Val(flxBankPayHist.TextMatrix(flxBankPayHist.row, 8)) + _
                       Val(flxBankPayHist.TextMatrix(flxBankPayHist.row, 9)), "0.00")
   End With

   frmBankTranEdit.Show
   Me.Enabled = False
End Sub

Private Sub cmdAddTrans_Click()
   If optBankReceipt.Value Or optBankPayment.Value Then
      Load frmBankTranEdit
      frmBankTranEdit.FrmBankTranEdit_CALLING_FROM = frmBankTransactions.Name

      frmBankTranEdit.Caption = "Add " & IIf(optBankReceipt.Value, optBankReceipt.Caption, optBankPayment.Caption)

      frmBankTranEdit.Show
   End If
   If optBankTransfer.Value Then
      Load frmBankTransfer
      frmBankTransfer.FrmBankTransfer_CALLING_FROM = "frmBankTransactions"

      frmBankTransfer.Show
   End If
   fraAddTrans.Visible = False
   tabPayment.Enabled = True
   frmBankTransactions.Enabled = False
End Sub

Private Sub cmdEditBk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub cmdNewBk_Click()
   fraAddTrans.Left = fraButtons.Left + cmdNewBk.Left + cmdNewBk.Width + tabPayment.Left
   fraAddTrans.Top = fraButtons.Top + cmdNewBk.Top + tabPayment.Top
   fraAddTrans.Visible = True
   fraAddTrans.ZOrder 0

   tabPayment.Enabled = False
   cmdAddTrans.SetFocus
   bEditMode = False
'  the following code display the options as copy/copy reverese
'   frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + fraButtons.Top + cmdNewBk.Top + frmBankTransactions.Top + tabPayment.Top + 1160
'   frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + fraButtons.Left + frmBankTransactions.Left + tabPayment.Left + cmdNewBk.Left + 80
'   frmPopUpMenu.CallingFrom "Bank"
'   frmPopUpMenu.Show
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
   If MsgBox("Do you wish to post these transactions to history?", vbQuestion + vbYesNo, "Bank Transactions") = vbNo Then Exit Sub

   Dim szBank As String
   Dim szSQL   As String
   Dim adoConn As New ADODB.Connection

   szBank = SelectedDemandID()

   If szBank = "" Then
      ShowMsgInTaskBar "No record has been seleted to post to history", "Y", "N"
      Exit Sub
   End If

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "UPDATE tlbBankPayment " & _
           "SET    UPDATE_SAGE = TRUE " & _
           "WHERE  UPDATE_SAGE = FALSE AND " & _
                  "MY_ID IN (" & szBank & ");"

   adoConn.Execute szSQL

    Call LoadFlxBankPay(adoConn, "")
    Call LoadflxBankPayHist(adoConn, "")

   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Transactions have been posted to history successfully.", "Y", "P"
End Sub

Private Function SelectedDemandID() As String
   Dim i As Integer

   SelectedDemandID = ""
   For i = 1 To flxBankPay.Rows - 1
      If flxBankPay.TextMatrix(i, 0) = "X" Then
         SelectedDemandID = SelectedDemandID & "'" & flxBankPay.TextMatrix(i, 10) & "'"
         SelectedDemandID = SelectedDemandID & ","
      End If
   Next i
   If SelectedDemandID = "" Then Exit Function

   SelectedDemandID = Left(SelectedDemandID, Len(SelectedDemandID) - 1)
End Function

Private Sub cmdPrintList_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\BankTransactions.rpt")
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
   Dim adoConn As New ADODB.Connection
   Dim strPartSql As String
   strPartSql = SelPurHistory
   On Error GoTo Catch_Error
   If strPartSql = "" Then Exit Sub
   If MsgBox("Are you sure to reverse back selected bank transaction from the history?", vbQuestion + vbYesNo, "Transaction History") = vbNo Then Exit Sub
   'szPurID = SelPurHistory()

   adoConn.Open getConnectionString
'
'   szSQL = "UPDATE tblPurInv " & _
'           "SET    History = FALSE " & _
'           "WHERE  SlNumber IN (" & SelPurHistory & ") ; "

'           AND " & _
'                  "  DemandID NOT IN (" & _
'                  "     SELECT DemandRef " & _
'                  "     From tlbReceipt " & _
'                  "     WHERE Type = 1 AND Amount > OSAmount);"
'Debug.Print szSQL
  ' adoConn.Execute szSQL

'   LoadFlxPurchHistory adoConn
'   LoadFlxPurchase adoConn
    szSQL = "UPDATE tlbBankPayment " & _
           "SET    UPDATE_SAGE = FALSE " & _
           "WHERE  " & _
                  "MY_ID IN (" & strPartSql & ");"

   adoConn.Execute szSQL

   Call LoadFlxBankPay(adoConn, "")
    Call LoadflxBankPayHist(adoConn, "")

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

Catch_Error:
   MsgBox "Select a bank transaction to reverese.", vbCritical + vbOKOnly, "Warning"
End Sub
Private Function SelPurHistory() As String
   Dim i As Integer

   SelPurHistory = ""
   For i = 1 To flxBankPayHist.Rows - 1
      If flxBankPayHist.TextMatrix(i, 0) = "X" Then
         'SelPurHistory = SelPurHistory & CStr(Mid(flxBankPayHist.TextMatrix(i, 1), 3, Len(flxBankPayHist.TextMatrix(i, 1))))
         SelPurHistory = SelPurHistory & "'" & CStr(flxBankPayHist.TextMatrix(i, 12)) & "'"
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
         'Call LoadFlxBankPay(adoconn, "")
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

Private Sub flxBankPay_Click()
   If flxBankPay.TextMatrix(flxBankPay.row, 0) = "" And flxBankPay.TextMatrix(flxBankPay.row, 1) <> "" Then
      flxBankPay.TextMatrix(flxBankPay.row, 0) = "X"
   Else
      flxBankPay.TextMatrix(flxBankPay.row, 0) = ""
   End If
End Sub

Public Sub CopyTransaction()
   Dim szStr As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   On Error GoTo ErrHandler
'      connect to database
   adoConn.Open getConnectionString

   szStr = "SELECT BP.* " & _
           "FROM tlbBankPayment AS BP;"
'Debug.Print szStr
   adoRST.Open szStr, adoConn, adOpenDynamic, adLockOptimistic
   With adoRST
      If .EOF Then
         GoTo ErrHandler
      Else
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = flxBankPay.TextMatrix(flxBankPay.row, 13)
         .Fields.Item("BANK_AC").Value = flxBankPay.TextMatrix(flxBankPay.row, 14)
         .Fields.Item("PropertyID").Value = flxBankPay.TextMatrix(flxBankPay.row, 11)
         .Fields.Item("UNIT_ID").Value = flxBankPay.TextMatrix(flxBankPay.row, 12)
         .Fields.Item("DESCRIPTION").Value = flxBankPay.TextMatrix(flxBankPay.row, 8)
         .Fields.Item("PROJ_REF").Value = flxBankPay.TextMatrix(flxBankPay.row, 7)
         .Fields.Item("NOMINAL_CODE").Value = flxBankPay.TextMatrix(flxBankPay.row, 5)
         .Fields.Item("DEPT_ID").Value = flxBankPay.TextMatrix(flxBankPay.row, 6)
         .Fields.Item("TRAN_DATE").Value = Format(Now, "dd mmmm yyyy")
         .Fields.Item("NET_AMOUNT").Value = CCur(flxBankPay.TextMatrix(flxBankPay.row, 15))
         .Fields.Item("TAX_CODE").Value = flxBankPay.TextMatrix(flxBankPay.row, 18)
         .Fields.Item("VAT").Value = flxBankPay.TextMatrix(flxBankPay.row, 16)

         If Left(flxBankPay.TextMatrix(flxBankPay.row, 1), 2) = "BR" Then
            .Fields.Item("TransactionType").Value = 12
            .Fields.Item("TRANS").Value = "BR"
         End If
         If Left(flxBankPay.TextMatrix(flxBankPay.row, 1), 2) = "BP" Then
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
         End If

         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Update
         .Close
      End If
   End With

   ShowMsgInTaskBar "The Transaction has been copied sucessfully.", "Y", "P"

   Set adoRST = Nothing

   flxBankPay.Clear
   flxBankPay.Rows = 2

   Call LoadFlxBankPay(adoConn, "")

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

ErrHandler:
   MsgBox "System could not update the record.", vbExclamation + vbOKOnly, "Edit Bank Transactions"

   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub CopyRevTransaction()
   Dim szStr As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   On Error GoTo ErrHandler
'      connect to database
   adoConn.Open getConnectionString

   szStr = "SELECT BP.* " & _
           "FROM tlbBankPayment AS BP;"
'Debug.Print szStr
   adoRST.Open szStr, adoConn, adOpenDynamic, adLockOptimistic
   With adoRST
      If .EOF Then
         GoTo ErrHandler
      Else
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = flxBankPay.TextMatrix(flxBankPay.row, 13)
         .Fields.Item("BANK_AC").Value = flxBankPay.TextMatrix(flxBankPay.row, 14)
         .Fields.Item("PropertyID").Value = flxBankPay.TextMatrix(flxBankPay.row, 11)
         .Fields.Item("UNIT_ID").Value = flxBankPay.TextMatrix(flxBankPay.row, 12)
         .Fields.Item("DESCRIPTION").Value = flxBankPay.TextMatrix(flxBankPay.row, 8)
         .Fields.Item("PROJ_REF").Value = flxBankPay.TextMatrix(flxBankPay.row, 7)
         .Fields.Item("NOMINAL_CODE").Value = flxBankPay.TextMatrix(flxBankPay.row, 5)
         .Fields.Item("DEPT_ID").Value = flxBankPay.TextMatrix(flxBankPay.row, 6)
         .Fields.Item("TRAN_DATE").Value = Format(Now, "dd mmmm yyyy")
         .Fields.Item("NET_AMOUNT").Value = CCur(flxBankPay.TextMatrix(flxBankPay.row, 15))
         .Fields.Item("TAX_CODE").Value = flxBankPay.TextMatrix(flxBankPay.row, 18)
         .Fields.Item("VAT").Value = flxBankPay.TextMatrix(flxBankPay.row, 16)

         If Left(flxBankPay.TextMatrix(flxBankPay.row, 1), 2) = "BR" Then
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
         End If
         If Left(flxBankPay.TextMatrix(flxBankPay.row, 1), 2) = "BP" Then
            .Fields.Item("TransactionType").Value = 12
            .Fields.Item("TRANS").Value = "BR"
         End If

         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoConn)
         .Update
         .Close
      End If
   End With

   ShowMsgInTaskBar "The Transaction has been copy reveresed sucessfully.", "Y", "P"

   Set adoRST = Nothing

   flxBankPay.Clear
   flxBankPay.Rows = 2
      
   Call LoadFlxBankPay(adoConn, "")

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

ErrHandler:
   MsgBox "System could not update the record.", vbExclamation + vbOKOnly, "Edit Bank Transactions"

   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxBankPay_DblClick()
    cmdEditBk_Click
End Sub
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
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

   
        adoConn.Open getConnectionString
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
               
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub flxBankPay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow

   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub flxBankPayHistHist_Click()
    If flxBankPayHist.TextMatrix(flxBankPayHist.row, 0) = "" And flxBankPayHist.TextMatrix(flxBankPayHist.row, 1) <> "" Then
      flxBankPayHist.TextMatrix(flxBankPayHist.row, 0) = "X"
   Else
      flxBankPayHist.TextMatrix(flxBankPayHist.row, 0) = ""
   End If
End Sub

Private Sub flxBankPayHist_Click()
     If flxBankPayHist.TextMatrix(flxBankPayHist.row, 0) = "" And flxBankPayHist.TextMatrix(flxBankPayHist.row, 1) <> "" Then
      flxBankPayHist.TextMatrix(flxBankPayHist.row, 0) = "X"
   Else
      flxBankPayHist.TextMatrix(flxBankPayHist.row, 0) = ""
   End If
End Sub

'Private Sub flxBankPayHist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Me.MousePointer = vbArrow
'End Sub

Private Sub flxClient_Click()
    tabPayment.Enabled = True
    Dim adoConn As New ADODB.Connection
     adoConn.Open getConnectionString
    If sTextBox = "1" Then
            txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
            txtPoperty.Tag = "ALL"
            txtPoperty.text = "ALL"
            ConfigFlxBankPay
            LoadFlxBankPay1 adoConn, 0, "ASC"
            cmdProperty.SetFocus
    Else
            txtPoperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtPoperty.text = flxClient.TextMatrix(flxClient.row, 2)
            ConfigFlxBankPay
            LoadFlxBankPay1 adoConn, 0, "ASC"
            txtDateFrom.SetFocus
    End If
    adoConn.Close
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
   Me.Width = 16665
   tabPayment.Tab = 0
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   UserSessionID = GetTimeStamp
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   ConfigFlxBankPay
   Call LoadFlxBankPay(adoConn, "")
   ConfigflxBankPayHist
   Call LoadflxBankPayHist(adoConn, "")
   'PrepareList adoConn, cmbClient, cmbProperty
   'LoadBankAccountInCombo adoConn
   
   txtTransTypeFilter.Tag = "ALL"
   txtTransTypeFilter.text = "ALL"
   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub
Private Sub LoadBankAccount()
   ' Error Handler
   'On Error GoTo Error_Handler
   configGridBankCode

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   If txtClientList.text = "ALL" Then
        rRow = 1
        gridBankCode.TextMatrix(rRow, 1) = "ALL"
        gridBankCode.TextMatrix(rRow, 2) = "ALL"
        Exit Sub
   End If
         
   adoConn.Open getConnectionString

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = '" & txtClientList.Tag & "' AND " & _
               "NominalLedger.ClientID = '" & txtClientList.Tag & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      MsgBox "Please setup bank account for this client : '" & txtClientList.text & "'", vbInformation, "Global bank account"
      picBankCode.Visible = False
   Else
      gridBankCode.Rows = adoRST.RecordCount + 2
      rRow = 1
        gridBankCode.TextMatrix(rRow, 1) = "ALL"
        gridBankCode.TextMatrix(rRow, 2) = "ALL"
      rRow = 2
      While Not adoRST.EOF
         gridBankCode.TextMatrix(rRow, 1) = adoRST.Fields.Item("BNC").Value
         gridBankCode.TextMatrix(rRow, 2) = adoRST.Fields.Item("BNN").Value
         rRow = rRow + 1
         adoRST.MoveNext
      Wend
       picBankCode.Visible = True
       gridBankCode.row = 1
   End If

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "Prestige Database Error: ", vbExclamation, "Load Bank Account in Demand"

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoConn = Nothing
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

Public Sub LoadflxBankPayHist(adoConn As ADODB.Connection, Filter As String)
   Dim iRow As Integer
   Dim szStr As String
   Dim tempstr As String
   Dim adoRST As New ADODB.Recordset

'   szStr = "SELECT BP.*, F.FundName, P.PropertyName,C.ClientName " & _
'           "FROM ((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = F.FundID) INNER JOIN CLIENT C ON C.ClientID=BP.ClientID)  LEFT JOIN " & _
'                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
'           "WHERE BP.DEPT_ID = F.FundID AND " & _
'               "BP.UPDATE_SAGE = TRUE " & _
'           "ORDER BY BP.TRANS, CLNG(BP.TRAN_ID);"
'  szStr = "SELECT BP.*, F.FundName, P.PropertyName,C.ClientName,NL.Name,(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1 " & _
'           "FROM (((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = F.FundID)  " & _
'           "INNER JOIN CLIENT C ON C.ClientID=BP.ClientID) INNER JOIN NominalLedger NL ON NL.ClientID=BP.ClientID AND NL.Code=BP.NOMINAL_CODE) LEFT JOIN " & _
'                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
'           "WHERE BP.DEPT_ID = F.FundID AND " & _
'               "BP.UPDATE_SAGE = TRUE " & _
'           "ORDER BY BP.TRANS, CLNG(BP.TRAN_ID);" 'UPDATE_SAGE = TRUE means history is true
    szStr = "SELECT TRANS & TRAN_ID as INVNO, BP.*, F.FundName, P.PropertyName,C.ClientName,NL.Name,(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1 " & _
           "FROM (((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = F.FundID)  " & _
           "INNER JOIN CLIENT C ON C.ClientID=BP.ClientID) INNER JOIN NominalLedger NL ON NL.ClientID=BP.ClientID AND NL.Code=BP.NOMINAL_CODE) LEFT JOIN " & _
                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
           "WHERE BP.DEPT_ID = F.FundID AND " & _
               "BP.UPDATE_SAGE = TRUE " & _
            "ORDER BY  TRANS,CLNG(TRAN_ID) desc" 'UPDATE_SAGE = TRUE means history is true
    If Filter = "3" Then
        szStr = "SELECT TRANS & TRAN_ID as INVNO, BP.*, F.FundName, P.PropertyName,C.ClientName,NL.Name,(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1 " & _
           "FROM (((tlbBankPayment AS BP INNER JOIN Fund AS F ON BP.DEPT_ID = F.FundID)  " & _
           "INNER JOIN CLIENT C ON C.ClientID=BP.ClientID) INNER JOIN NominalLedger NL ON NL.ClientID=BP.ClientID AND NL.Code=BP.NOMINAL_CODE) LEFT JOIN " & _
                 "Property AS P ON BP.PropertyID = P.PropertyID " & _
           "WHERE BP.DEPT_ID = F.FundID AND " & _
               "BP.UPDATE_SAGE = TRUE AND BP.TRAN_DATE >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND " & _
                    "BP.TRAN_DATE <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "# " & _
            "ORDER BY  TRANS,CLNG(TRAN_ID) desc"
    End If
    
   adoRST.Open szStr, adoConn, adOpenDynamic, adLockPessimistic
'Debug.Print szStr
'  szHeader$ = "<No|<Bank|<Date|<Property|<NC|<Ref|<Fund|<Details|>Net|>VAT|>Total|ID|Client|PropID|FundID|Recon|Unit|TC"
    
    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            adoRST.Filter = "INvNo Like '%" & tempstr & "%'"
        End If
    End If

 
    
    
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            If tabPayment.Tab = 0 Then
                adoRST.Filter = "Details Like '%" & tempstr & "%'"
            Else
                adoRST.Filter = "Description Like '%" & tempstr & "%'"
            End If
        End If
    End If
   flxBankPayHist.Clear
   flxBankPayHist.Rows = 2
   iRow = 1
   While Not adoRST.EOF
      flxBankPayHist.TextMatrix(iRow, 0) = ""
      flxBankPayHist.TextMatrix(iRow, 1) = adoRST!InvNo 'adoRst!TRANS & adoRst!TRAN_ID
      flxBankPayHist.TextMatrix(iRow, 2) = adoRST!BANK_AC
      flxBankPayHist.TextMatrix(iRow, 3) = adoRST!TRAN_DATE
      flxBankPayHist.TextMatrix(iRow, 4) = IIf(IsNull(adoRST!PropertyName), "", adoRST!PropertyName)
      flxBankPayHist.TextMatrix(iRow, 5) = IIf(IsNull(adoRST!Nominal_code), "", adoRST!Nominal_code)
      flxBankPayHist.TextMatrix(iRow, 6) = IIf(IsNull(adoRST!PROJ_REF), "", adoRST!PROJ_REF)
      flxBankPayHist.TextMatrix(iRow, 7) = IIf(IsNull(adoRST!FundName), "", adoRST!FundName)
      flxBankPayHist.TextMatrix(iRow, 8) = IIf(IsNull(adoRST!description), "", adoRST!description)
      flxBankPayHist.TextMatrix(iRow, 9) = Format(IIf(IsNull(adoRST!NET_AMOUNT), "", adoRST!NET_AMOUNT), "0.00")
      flxBankPayHist.TextMatrix(iRow, 10) = Format(IIf(IsNull(adoRST!vat), "0", adoRST!vat), "0.00")
      flxBankPayHist.TextMatrix(iRow, 11) = Format(Val(flxBankPayHist.TextMatrix(iRow, 8)) + _
                                          Val(flxBankPayHist.TextMatrix(iRow, 9)), "0.00")
      flxBankPayHist.TextMatrix(iRow, 12) = adoRST!My_ID
      flxBankPayHist.TextMatrix(iRow, 13) = IIf(IsNull(adoRST!ClientID), "", adoRST!ClientID)
      flxBankPayHist.TextMatrix(iRow, 14) = IIf(IsNull(adoRST!propertyID), "", adoRST!propertyID)
      flxBankPayHist.TextMatrix(iRow, 15) = adoRST!DEPT_ID
      flxBankPayHist.TextMatrix(iRow, 16) = IIf(IsNull(adoRST!ReconNow), "N", "Y")
      flxBankPayHist.TextMatrix(iRow, 17) = IIf(IsNull(adoRST!UNIT_ID), "", adoRST!UNIT_ID)
      flxBankPayHist.TextMatrix(iRow, 18) = IIf(IsNull(adoRST!TAX_CODE), "", adoRST!TAX_CODE)
      flxBankPayHist.TextMatrix(iRow, 19) = IIf(IsNull(adoRST!ClientName), "", adoRST!ClientName)
      flxBankPayHist.TextMatrix(iRow, 20) = IIf(IsNull(adoRST!Name), "", adoRST!Name) 'Nominal Bank Account Name
       flxBankPayHist.TextMatrix(iRow, 21) = IIf(IsNull(adoRST!Name1), "", adoRST!Name1) 'Nominal Account Name
      adoRST.MoveNext
      If Not adoRST.EOF Then flxBankPayHist.AddItem ""
      iRow = iRow + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ConfigflxBankPayHist()
   Dim szHeader As String, iCol As Integer

   szHeader$ = "<No|<Bank|<Date|<Property|<NC|<Ref|<Fund|<Details|>Net|>VAT|>Total|ID|Client|PropID|FundID|Recon|Unit|TC"

   flxBankPayHist.Clear
   flxBankPayHist.Rows = 2
   flxBankPayHist.Cols = 22

   flxBankPayHist.RowHeight(0) = 0
   flxBankPayHist.FormatString = szHeader$
   flxBankPayHist.ColWidth(0) = 300
   For iCol = 1 To flxBankPayHist.Cols - 13
      'flxBankPayHist.ColWidth(iCol - 1) = lblBankHist(iCol).Left - lblBankHist(iCol - 1).Left
     flxBankPayHist.ColWidth(iCol) = lblBankHist(iCol).Left - lblBankHist(iCol - 1).Left
   Next iCol

   'flxBankPayHist.ColWidth(10) = flxBankPayHist.Width + flxBankPayHist.Left - lblBankHist(10).Left - 300
   flxBankPayHist.ColWidth(10) = 700
   flxBankPayHist.ColWidth(11) = 800 ' flxBankPayHist.Width + flxBankPayHist.Left - lblBankHist(10).Left - 300
   'flxBankPayHist.ColWidth(11) = 0         'ID field
   flxBankPayHist.ColWidth(12) = 0         'Client Id
   flxBankPayHist.ColWidth(13) = 0         'Porperty ID
   flxBankPayHist.ColWidth(14) = 0         'FundID
   flxBankPayHist.ColWidth(15) = 0         'Reconciliation Y/N
   flxBankPayHist.ColWidth(16) = 0         'Unit
   flxBankPayHist.ColWidth(17) = 0         'Tax Code
   flxBankPayHist.ColWidth(18) = 0         'Client Name
   flxBankPayHist.ColWidth(19) = 0         'Bank Nominal Name
   flxBankPayHist.ColWidth(20) = 0         'Account Nominal Name
   flxBankPayHist.ColWidth(21) = 0         'Account Nominal Name
End Sub

Private Sub ConfigFlxBankPay()
   Dim szHeader As String, iCol As Integer

   flxBankPay.Clear
   flxBankPay.Rows = 2
   'flxBankPay.Cols = 25
   'adding 4 more col
   flxBankPay.Cols = 29
   flxBankPay.RowHeight(0) = 0

   szHeader$ = "Empty|<SL No|<Date|<Bank|<PropName" & _
               "|<NC|<Fund|<Ref|<Desc|>Amount|TranID|PD"
   flxBankPay.FormatString = szHeader$

   flxBankPay.ColWidth(0) = lblBankRec(0).Left - flxBankPay.Left

   For iCol = 1 To 8 'flxBankPay.Cols - 3 when cols was 10
      flxBankPay.ColWidth(iCol) = lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left
   Next iCol

   flxBankPay.ColWidth(iCol) = flxBankPay.Width + flxBankPay.Left - lblBankRec(iCol - 1).Left - 340
   For iCol = iCol + 1 To flxBankPay.Cols - 1
      flxBankPay.ColWidth(iCol) = 0
   Next iCol
   
End Sub
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
   selRow = flxBankPay.row
   selcol = flxBankPay.col
   
   
   adoConn.Open getConnectionString
   
'   ' I am doing some test here
'   ' on loading time full table vs selected row
'
'    szSQL = " SELECT BP.TransactionID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment BP WHERE MYid= '" & flxBankPay.TextMatrix(flxBankPay.row, 10) & "'"
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
   szSQL = " SELECT BP.TransactionID,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment BP WHERE MYid= '" & flxBankPay.TextMatrix(flxBankPay.row, 10) & "'"
   
   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

  'locking status show for current row
   If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 Then   'szSQL <> UserSessionID shall be always true bcoz PI is generating only one session thrgh out this module.
                flxBankPay.col = 0
                flxBankPay.row = flxBankPay.row
                flxBankPay.CellBackColor = vbRed
                InstantLockingCheck = False
                colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & IIf(IsNull(adoPay("TransactionID").Value), "", adoPay("TransactionID").Value) & ","
            Else 'lock for this user
                        flxBankPay.col = 0
                        i = flxBankPay.row
                        flxBankPay.CellBackColor = vbWhite
    '                    adoconn.Execute "Update tlbPayment Set  DateTimeStamp='" & Now & "',Module='Purchase Invoice',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
    '                    "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where tlbPayment.PI = '" & flxBankPay.TextMatrix(iPIEdit, 0) & "'"
                        'Need to clear the locking flag
                        flxBankPay.TextMatrix(i, 25) = ""
                        flxBankPay.TextMatrix(i, 26) = ""
                        flxBankPay.TextMatrix(i, 27) = ""
                        flxBankPay.TextMatrix(i, 28) = ""
            End If
           
   End If
   'second part instant unlock
       If Len(colTransactionIDOtherPIGrid) > 0 Then
            szSQL = "SELECT TransactionID,UserSessionID,WindowsUserName,MachineName,Module,ClientID " & _
                 "FROM tlbBankPayment where (isnull(UserSessionID) OR UserSessionID='') " & _
                 " AND TransactionID in (" & colTransactionIDOtherPIGrid & ") order by 2,3 Desc;"
            rsLockDialog.Open szSQL, adoConn, adOpenStatic, adLockReadOnly 'Selecting those transaction which has been unlocked in the background with out knowing this form
             While Not rsLockDialog.EOF
                      flxBankPay.col = 0
                      For j = 1 To flxBankPay.Rows - 1
                          If flxBankPay.TextMatrix(j, 10) = rsLockDialog("transactionID").Value And i <> j Then 'no need to update row of first part check
                                flxBankPay.row = j
                                flxBankPay.CellBackColor = vbWhite
                          End If
                       Next j
                    rsLockDialog.MoveNext
              Wend
        End If
        'second part ends here
        
        flxBankPay.col = selcol
        flxBankPay.row = selRow
        
        
        
        
   adoPay.Close
   adoConn.Close
   flxBankPay.row = selRow
   flxBankPay.col = selcol
   Set adoPay = Nothing
   Set adoConn = Nothing
End Function
Private Function IsPossible2Edit() As Boolean
  Dim adoPay As New ADODB.Recordset
   Dim adoRST As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer
   Dim strTemp As String

   adoConn.Open getConnectionString
   szSQL = " SELECT BP.TRANS,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module,BP.ClientID FROM tlbBankPayment BP WHERE MY_ID= '" & flxBankPay.TextMatrix(flxBankPay.row, 10) & "'"
   adoPay.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoPay.EOF Then
            szSQL = IIf(IsNull(adoPay("UserSessionID").Value), "", adoPay("UserSessionID").Value)
            If Len(szSQL) > 0 Then  'szSQL <> UserSessionID shall be always true bcoz PI is generating only one seeion through out this module.
                flxBankPay.col = 0
                'flxBankPay.row = iPIEdit
                flxBankPay.CellBackColor = vbRed
                strTemp = IIf(adoPay("TRANS").Value = "BP", "Bank Payment", "Bank Receipt")
                MsgBox "The selected " & strTemp & " is currently locked by '" & IIf(IsNull(adoPay("WindowsUserName").Value), "", adoPay("WindowsUserName").Value) & _
                "' on '" & IIf(IsNull(adoPay("MachineName").Value), "", adoPay("MachineName").Value) & "' in the '" & IIf(IsNull(adoPay("Module").Value), "", adoPay("Module").Value) & "'" & vbCrLf & "" & _
                        "screen for the Client '" & IIf(IsNull(adoPay("ClientID").Value), "", adoPay("ClientID").Value) & "' and cannot be edited. Please wait until it is released.", vbInformation, "Warning"
                IsPossible2Edit = False
            Else 'lock row for this user in database
                flxBankPay.col = 0
                IsPossible2Edit = True
                flxBankPay.CellBackColor = vbWhite
            End If
   End If
   adoPay.Close
   Set adoPay = Nothing
   
   If IsPossible2Edit = True Then 'the reason I need to lock it here because it has passed all the tests here to open PI and now I can lock
             adoConn.Execute "Update tlbBankPayment Set  DateTimeStamp='" & Now & "',Module='Bank Payment',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
                "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where MY_ID = '" & flxBankPay.TextMatrix(flxBankPay.row, 10) & "'"
   End If
   adoConn.Close
   Set adoConn = Nothing
End Function
Private Sub cmdEditBk_Click()
   Dim iCol As Integer
   Dim i As Integer
   Dim rCount As Integer
   Dim iIncDec As Integer
   For rCount = 1 To flxBankPay.Rows - 1
        If flxBankPay.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            iBankPayRow = rCount
        End If
   Next
   If iIncDec <> 1 Then 'bebugger shall not go further if you do not select one row
      MsgBox "Please select one transaction only.", vbInformation + vbOKOnly, "Transaction Selection"
      chkSelAll.Value = 0
      For i = 1 To flxBankPay.Rows - 1
           If flxBankPay.TextMatrix(i, 0) = "X" Then
                flxBankPay.TextMatrix(i, 0) = ""
           End If
      Next i
      Exit Sub
   End If
   flxBankPay.row = iBankPayRow
   If flxBankPay.TextMatrix(iBankPayRow, 17) = "Y" Then
            'added by anol 01 Sep 2016
            flxBankPay.TextMatrix(iBankPayRow, 0) = ""
            If flxBankPay.row > 0 Then
                If flxBankPay.TextMatrix(flxBankPay.row, 0) = "" Then
                         For iCol = 1 To flxBankPay.Cols - 1
                            flxBankPay.col = iCol
                            flxBankPay.CellBackColor = RGB(174, 199, 200)
                         Next iCol
                End If
            End If
            MsgBox "The transaction has been bank reconciled.", vbInformation, "Warning"
            Exit Sub
   End If
   'Check lock if you can edit
   If Not IsPossible2Edit Then
        Exit Sub
   End If
   bEditMode = True
   Load frmBankTranEdit
   frmBankTranEdit.Caption = "Edit Bank Transaction - " & flxBankPay.TextMatrix(flxBankPay.row, 1)
   frmBankTranEdit.FrmBankTranEdit_CALLING_FROM = Me.Name
'   frmBankTranEdit.FrmBankTranEdit_CALLING_MODE = "Edit"

   With frmBankTranEdit
      .szTransID = flxBankPay.TextMatrix(flxBankPay.row, 10)
      .txtClientList.Tag = flxBankPay.TextMatrix(flxBankPay.row, 13)
      .txtClientList.text = flxBankPay.TextMatrix(flxBankPay.row, 20)
      'Modified by anol 20160512
      '.cboBC.Value = flxBankPay.TextMatrix(flxBankPay.row, 14)
      .txtBankName.text = flxBankPay.TextMatrix(flxBankPay.row, 21)
      .txtBankCode.text = flxBankPay.TextMatrix(flxBankPay.row, 14)
      .txtProperty.Tag = flxBankPay.TextMatrix(flxBankPay.row, 11)
      .txtProperty.text = flxBankPay.TextMatrix(flxBankPay.row, 4)
      .txtUnit.Tag = flxBankPay.TextMatrix(flxBankPay.row, 12)
      .txtUnit.text = flxBankPay.TextMatrix(flxBankPay.row, 22)
      'Resolved by BOSL
      'Issue 452
      'modified by anol 13 Aug 2014
      '.cboUnit1.text = .cboUnit.text
      .txtDetails.text = flxBankPay.TextMatrix(flxBankPay.row, 8)
      .txtReference.text = flxBankPay.TextMatrix(flxBankPay.row, 7)
      .txtNC.Tag = flxBankPay.TextMatrix(flxBankPay.row, 5)
      .txtNC.text = flxBankPay.TextMatrix(flxBankPay.row, 24)
      .txtFund.Tag = flxBankPay.TextMatrix(flxBankPay.row, 6)
      .txtFund.text = flxBankPay.TextMatrix(flxBankPay.row, 23)
      .txtDate.text = flxBankPay.TextMatrix(flxBankPay.row, 2)
      .txtNet.text = Format(flxBankPay.TextMatrix(flxBankPay.row, 15), "0.00")
      '.cboVat.Value = flxBankPay.TextMatrix(flxBankPay.row, 18)
      .Label1(24).Caption = flxBankPay.TextMatrix(flxBankPay.row, 18)
       'Resolved by BOSL
      'Issue 463
      'modified by anol 13 Aug 2014
      .txtVat_.text = Format(Val(flxBankPay.TextMatrix(flxBankPay.row, 16)), "0.00")
      'end fo modification
      .txtTotal.text = Format(Val(flxBankPay.TextMatrix(flxBankPay.row, 9)), "0.00")
      .lblPostingDate.ToolTipText = Format(flxBankPay.TextMatrix(flxBankPay.row, 19), "dd/mm/yyyy")
      If .txtDetails.text = "BANK TRANSFER" Then
            .txtDetails.Locked = True
            .cmdNC.Enabled = False
            .Label6(12).Visible = False
            .Label1(24).Visible = False
            .cmdVATCode.Visible = False
            .txtVat_.Visible = False

      Else
            .txtDetails.Locked = False
            .cmdNC.Enabled = True
            .Label6(12).Visible = True
            .Label1(24).Visible = True
            .cmdVATCode.Visible = True
            .txtVat_.Visible = True
      End If
                
                
   End With
    'added by anol 01 Sep 2016
        flxBankPay.TextMatrix(flxBankPay.row, 0) = ""
'        flxBankPay.CellBackColor = vbCyan 'RGB(179, 233, 174)
            If flxBankPay.row > 0 Then
            If flxBankPay.TextMatrix(flxBankPay.row, 0) = "" Then
                 'If flxBankPay.CellBackColor = RGB(174, 179, 233) Then
                     For iCol = 1 To flxBankPay.Cols - 1
                        flxBankPay.col = iCol
                        flxBankPay.CellBackColor = RGB(174, 199, 200)
                     Next iCol
                ' End If
            End If
         End If
   frmBankTranEdit.Show
   Me.Enabled = False
End Sub

Public Sub LoadFlxBankPay(adoConn As ADODB.Connection, Filter As String)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRST As New ADODB.Recordset
   Dim tempstr As String
   colTransactionIDOtherPIGrid = ""
'  Column Heading: "Empty|<SL No|<Date|<Bank|<PropName|<NC|<Fund|<Ref|<Desc|>Amount|TranID|PD"
''                     ^      ^      ^     ^      ^      ^     ^    ^     ^       ^     ^
'   szSQL = "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
'                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, BP.MY_ID, C.CLIENT_ID, " & _
'                  "TT.DESCRIPTION AS Type1, BP.PROJ_REF AS Rfn, BP.PropertyID, BP.UNIT_ID, BP.NOMINAL_CODE, " & _
'                  "BP.BANK_AC AS ACC, C.Bank_AC_Name, BP.DEPT_ID, BP.NET_AMOUNT, BP.VAT, BP.ReconNow, " & _
'                  "BP.TAX_CODE, P.PropertyName, BP.PostingDate " & _
'           "FROM ((tlbBankPayment AS BP INNER JOIN tlbTransactionTypes AS TT ON BP.TransactionType = TT.TYPE_ID) INNER JOIN " & _
'                  "tlbClientBanks AS C ON BP.BANK_AC = C.NominalCode) LEFT JOIN " & _
'                  "Property AS P ON BP.PropertyID = P.PropertyID " & _
'           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
'                  "C.CLIENT_ID = BP.ClientID AND " & _
'                  "BP.UPDATE_SAGE = FALSE " & _
'           "ORDER BY TRANS, CLNG(TRAN_ID)"
'added client table on query by anol 05122016
',(Select UnitName FROM UNITS where UnitNumber=UNIT_ID) AS UNITNAME
 szSQL = "SELECT MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)  & BP.TRAN_ID AS INvNo, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, BP.MY_ID, C.CLIENT_ID, " & _
                  "TT.DESCRIPTION AS Type1, BP.PROJ_REF AS Rfn, BP.PropertyID, BP.UNIT_ID, BP.NOMINAL_CODE, " & _
                  "BP.BANK_AC AS ACC, NL.Name AS NN, BP.DEPT_ID, BP.NET_AMOUNT, BP.VAT, BP.ReconNow, " & _
                  "BP.TAX_CODE, P.PropertyName, BP.PostingDate, D.ClientName,(Select UnitName FROM UNITS where " & _
                  "UnitNumber=BP.UNIT_ID) as UnitName,(Select FundName from FUND where CSTR(FundID)=BP.DEPT_ID) AS FUNDNAME," & _
                  "(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module " & _
           "FROM ((((tlbBankPayment AS BP INNER JOIN CLIENT D ON D.ClientID=BP.ClientID) INNER JOIN tlbTransactionTypes AS TT ON BP.TransactionType = TT.TYPE_ID)  INNER JOIN " & _
                  "tlbClientBanks AS C ON BP.BANK_AC = C.NominalCode) INNER JOIN NominalLedger NL ON NL.ClientID=C.CLIENT_ID AND NL.CODE=C.NominalCode) LEFT JOIN " & _
                  "Property AS P ON BP.PropertyID = P.PropertyID " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                  "C.CLIENT_ID = BP.ClientID AND " & _
                  "BP.UPDATE_SAGE = FALSE " & _
           "ORDER BY  CLNG(TRAN_ID) desc"
           '"ORDER BY  TRANS,CLNG(TRAN_ID) desc" we dont want to group by TRANS

'Debug.Print szSQL
'BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2
   If Filter = "3" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
'            szSQL = "SELECT S.*, C.ClientName, P.PropertyName " & _
'                    "FROM (NJ_Header AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID) " & _
'                    "LEFT JOIN Property AS P ON S.PropertyID = P.PropertyID " & _
'                    "WHERE NOT History AND S.TRAN_DATE >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND " & _
'                    "S.TRAN_DATE <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "#  order by RecordID DESC;"
                    
             szSQL = "SELECT MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)  & BP.TRAN_ID AS INvNo, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, BP.MY_ID, C.CLIENT_ID, " & _
                  "TT.DESCRIPTION AS Type1, BP.PROJ_REF AS Rfn, BP.PropertyID, BP.UNIT_ID, BP.NOMINAL_CODE, " & _
                  "BP.BANK_AC AS ACC, NL.Name AS NN, BP.DEPT_ID, BP.NET_AMOUNT, BP.VAT, BP.ReconNow, " & _
                  "BP.TAX_CODE, P.PropertyName, BP.PostingDate, D.ClientName,(Select UnitName FROM UNITS where " & _
                  "UnitNumber=BP.UNIT_ID) as UnitName,(Select FundName from FUND where CSTR(FundID)=BP.DEPT_ID) AS FUNDNAME," & _
                  "(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1,BP.UserSessionID,BP.WindowsUserName,BP.MachineName ,BP.Module  " & _
           "FROM ((((tlbBankPayment AS BP INNER JOIN CLIENT D ON D.ClientID=BP.ClientID) INNER JOIN tlbTransactionTypes AS TT ON BP.TransactionType = TT.TYPE_ID)  INNER JOIN " & _
                  "tlbClientBanks AS C ON BP.BANK_AC = C.NominalCode) INNER JOIN NominalLedger NL ON NL.ClientID=C.CLIENT_ID AND NL.CODE=C.NominalCode) LEFT JOIN " & _
                  "Property AS P ON BP.PropertyID = P.PropertyID " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                  "C.CLIENT_ID = BP.ClientID AND " & _
                  "BP.UPDATE_SAGE = FALSE AND BP.TRAN_DATE >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND " & _
                    "BP.TRAN_DATE <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "# " & _
           "ORDER BY  TRANS,CLNG(TRAN_ID) desc"
           
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
    End If
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly


    If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            adoRST.Filter = "INvNo Like '%" & tempstr & "%'"
        End If
    End If

 
    
    
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            adoRST.Filter = "Details Like '%" & tempstr & "%'"
        End If
    End If
   i = 1
   
    
    If adoRST.RecordCount = 0 Then
        flxBankPay.Clear
        flxBankPay.Rows = 2
    Else
        flxBankPay.Rows = adoRST.RecordCount + 1
    End If
    
    
   
   While Not adoRST.EOF
      flxBankPay.TextMatrix(i, 1) = adoRST.Fields.Item("INvNo").Value  'SL No
      'adoRst.Fields.Item("Type2").Value &  adoRst.Fields.Item("T_ID").Value                            'SL No
      flxBankPay.TextMatrix(i, 2) = adoRST.Fields.Item("RDate").Value                           'Date
      flxBankPay.TextMatrix(i, 3) = adoRST.Fields.Item("ACC").Value                             'Account Name
      flxBankPay.TextMatrix(i, 4) = IIf(IsNull(adoRST!PropertyName), "", adoRST!PropertyName)   'Property Name
      flxBankPay.TextMatrix(i, 5) = IIf(IsNull(adoRST!Nominal_code), "", adoRST!Nominal_code)   'NOMINAL_CODE
      flxBankPay.TextMatrix(i, 6) = adoRST.Fields.Item("DEPT_ID").Value                         'Fund
      flxBankPay.TextMatrix(i, 7) = adoRST.Fields.Item("Rfn").Value                             'Ref
      flxBankPay.TextMatrix(i, 8) = adoRST.Fields.Item("Details").Value                         'Desc
      flxBankPay.TextMatrix(i, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")          'Amt
      flxBankPay.TextMatrix(i, 10) = adoRST.Fields.Item("MY_ID").Value                          'TranID
      flxBankPay.TextMatrix(i, 11) = IIf(IsNull(adoRST!propertyID), "", adoRST!propertyID)      'PropertyID
      flxBankPay.TextMatrix(i, 12) = IIf(IsNull(adoRST!UNIT_ID), "", adoRST!UNIT_ID)            'UNIT_ID
      flxBankPay.TextMatrix(i, 13) = IIf(IsNull(adoRST!CLIENT_ID), "", adoRST!CLIENT_ID)        'CLIENT ID
      flxBankPay.TextMatrix(i, 14) = IIf(IsNull(adoRST!ACC), "", adoRST!ACC)                    'BANK_AC
      flxBankPay.TextMatrix(i, 15) = adoRST.Fields.Item("NET_AMOUNT").Value                     'NET_AMOUNT
      flxBankPay.TextMatrix(i, 16) = adoRST.Fields.Item("VAT").Value                            'VAT
      If adoRST.Fields.Item("INvNo").Value = "BR26" Then
        Debug.Print ""
      End If
      
      flxBankPay.TextMatrix(i, 17) = IIf(IsNull(adoRST!ReconNow) Or adoRST!ReconNow = "", "N", "Y")
      flxBankPay.TextMatrix(i, 18) = IIf(IsNull(adoRST!TAX_CODE), "", adoRST!TAX_CODE)
      flxBankPay.TextMatrix(i, 19) = IIf(IsNull(adoRST!postingDate), adoRST.Fields.Item("RDate").Value, adoRST!postingDate)    'Posting Date
      flxBankPay.TextMatrix(i, 20) = adoRST!ClientName    'CLIENT Name
      flxBankPay.TextMatrix(i, 21) = adoRST!NN    'Nominal Bank AC  Name
'      Dim rsUnit As New ADODB.Recordset
'      rsUnit.Open "Select UnitName FROM UNITS where UnitNumber='" & IIf(IsNull(adoRst!UNIT_ID), "", adoRst!UNIT_ID) & "'", adoConn, adOpenKeyset, adLockReadOnly
'      If Not rsUnit.EOF Then
'        flxBankPay.TextMatrix(i, 22) = "UNit Name" 'IIf(IsNull(rsUnit.Fields(0).Value), "", rsUnit.Fields(0).Value)
'      End If
        flxBankPay.TextMatrix(i, 22) = IIf(IsNull(adoRST!UnitName), "", adoRST!UnitName)
        flxBankPay.TextMatrix(i, 23) = IIf(IsNull(adoRST!FundName), "", adoRST!FundName)
        flxBankPay.TextMatrix(i, 24) = IIf(IsNull(adoRST!Name1), "", adoRST!Name1) 'Nominal account Name
        
        flxBankPay.TextMatrix(i, 25) = IIf(IsNull(adoRST!UserSessionID), "", adoRST!UserSessionID)
        flxBankPay.TextMatrix(i, 26) = IIf(IsNull(adoRST!WindowsUserName), "", adoRST!WindowsUserName)
        flxBankPay.TextMatrix(i, 27) = IIf(IsNull(adoRST!MachineName), "", adoRST!MachineName)
        flxBankPay.TextMatrix(i, 28) = IIf(IsNull(adoRST!Module), "", adoRST!Module)
        If flxBankPay.TextMatrix(i, 25) <> "" Then
                flxBankPay.row = i
                flxBankPay.col = 0
                flxBankPay.CellBackColor = vbRed
                colTransactionIDOtherPIGrid = colTransactionIDOtherPIGrid & adoRST.Fields.Item("MY_ID").Value & ","
        End If
'      rsUnit.Close
      adoRST.MoveNext
      'If Not adoRst.EOF Then flxBankPay.AddItem ""
      i = i + 1
   Wend
   If Len(colTransactionIDOtherPIGrid) > 0 Then
        colTransactionIDOtherPIGrid = Left(colTransactionIDOtherPIGrid, Len(colTransactionIDOtherPIGrid) - 1)
   End If
   adoRST.Close
   Set adoRST = Nothing
End Sub

Public Sub LoadFlxBankPaybyclient(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRST As New ADODB.Recordset

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
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   flxBankPay.Clear
   flxBankPay.Rows = 2
   While Not adoRST.EOF
      flxBankPay.TextMatrix(i, 1) = adoRST.Fields.Item("Type2").Value & _
                                    adoRST.Fields.Item("T_ID").Value                            'SL No
      flxBankPay.TextMatrix(i, 2) = adoRST.Fields.Item("RDate").Value                           'Date
      flxBankPay.TextMatrix(i, 3) = adoRST.Fields.Item("ACC").Value                             'Account Name
      flxBankPay.TextMatrix(i, 4) = IIf(IsNull(adoRST!PropertyName), "", adoRST!PropertyName)   'Property Name
      flxBankPay.TextMatrix(i, 5) = IIf(IsNull(adoRST!Nominal_code), "", adoRST!Nominal_code)   'NOMINAL_CODE
      flxBankPay.TextMatrix(i, 6) = adoRST.Fields.Item("DEPT_ID").Value                         'Fund
      flxBankPay.TextMatrix(i, 7) = adoRST.Fields.Item("Rfn").Value                             'Ref
      flxBankPay.TextMatrix(i, 8) = adoRST.Fields.Item("Details").Value                         'Desc
      flxBankPay.TextMatrix(i, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")          'Amt
      flxBankPay.TextMatrix(i, 10) = adoRST.Fields.Item("MY_ID").Value                          'TranID
      flxBankPay.TextMatrix(i, 11) = IIf(IsNull(adoRST!propertyID), "", adoRST!propertyID)      'PropertyID
      flxBankPay.TextMatrix(i, 12) = IIf(IsNull(adoRST!UNIT_ID), "", adoRST!UNIT_ID)            'UNIT_ID
      flxBankPay.TextMatrix(i, 13) = IIf(IsNull(adoRST!CLIENT_ID), "", adoRST!CLIENT_ID)        'CLIENT ID
      flxBankPay.TextMatrix(i, 14) = IIf(IsNull(adoRST!ACC), "", adoRST!ACC)                    'BANK_AC
      flxBankPay.TextMatrix(i, 15) = adoRST.Fields.Item("NET_AMOUNT").Value                     'NET_AMOUNT
      flxBankPay.TextMatrix(i, 16) = adoRST.Fields.Item("VAT").Value                            'VAT
      flxBankPay.TextMatrix(i, 17) = IIf(IsNull(adoRST!ReconNow), "N", "Y")
      flxBankPay.TextMatrix(i, 18) = IIf(IsNull(adoRST!TAX_CODE), "", adoRST!TAX_CODE)
      flxBankPay.TextMatrix(i, 19) = IIf(IsNull(adoRST!postingDate), adoRST.Fields.Item("RDate").Value, adoRST!postingDate)    'Posting Date

      adoRST.MoveNext
      If Not adoRST.EOF Then flxBankPay.AddItem ""
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing
End Sub
Public Sub LoadFlxBankPaybyProperty(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRST As New ADODB.Recordset

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
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   flxBankPay.Clear
   flxBankPay.Rows = 2
   While Not adoRST.EOF
      flxBankPay.TextMatrix(i, 1) = adoRST.Fields.Item("Type2").Value & _
                                    adoRST.Fields.Item("T_ID").Value                            'SL No
      flxBankPay.TextMatrix(i, 2) = adoRST.Fields.Item("RDate").Value                           'Date
      flxBankPay.TextMatrix(i, 3) = adoRST.Fields.Item("ACC").Value                             'Account Name
      flxBankPay.TextMatrix(i, 4) = IIf(IsNull(adoRST!PropertyName), "", adoRST!PropertyName)   'Property Name
      flxBankPay.TextMatrix(i, 5) = IIf(IsNull(adoRST!Nominal_code), "", adoRST!Nominal_code)   'NOMINAL_CODE
      flxBankPay.TextMatrix(i, 6) = adoRST.Fields.Item("DEPT_ID").Value                         'Fund
      flxBankPay.TextMatrix(i, 7) = adoRST.Fields.Item("Rfn").Value                             'Ref
      flxBankPay.TextMatrix(i, 8) = adoRST.Fields.Item("Details").Value                         'Desc
      flxBankPay.TextMatrix(i, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")          'Amt
      flxBankPay.TextMatrix(i, 10) = adoRST.Fields.Item("MY_ID").Value                          'TranID
      flxBankPay.TextMatrix(i, 11) = IIf(IsNull(adoRST!propertyID), "", adoRST!propertyID)      'PropertyID
      flxBankPay.TextMatrix(i, 12) = IIf(IsNull(adoRST!UNIT_ID), "", adoRST!UNIT_ID)            'UNIT_ID
      flxBankPay.TextMatrix(i, 13) = IIf(IsNull(adoRST!CLIENT_ID), "", adoRST!CLIENT_ID)        'CLIENT ID
      flxBankPay.TextMatrix(i, 14) = IIf(IsNull(adoRST!ACC), "", adoRST!ACC)                    'BANK_AC
      flxBankPay.TextMatrix(i, 15) = adoRST.Fields.Item("NET_AMOUNT").Value                     'NET_AMOUNT
      flxBankPay.TextMatrix(i, 16) = adoRST.Fields.Item("VAT").Value                            'VAT
      flxBankPay.TextMatrix(i, 17) = IIf(IsNull(adoRST!ReconNow), "N", "Y")
      flxBankPay.TextMatrix(i, 18) = IIf(IsNull(adoRST!TAX_CODE), "", adoRST!TAX_CODE)
      flxBankPay.TextMatrix(i, 19) = IIf(IsNull(adoRST!postingDate), adoRST.Fields.Item("RDate").Value, adoRST!postingDate)    'Posting Date

      adoRST.MoveNext
      If Not adoRST.EOF Then flxBankPay.AddItem ""
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing
End Sub
Private Sub LoadBankAccountInCombo(ByVal adoConn As ADODB.Connection)
   On Error GoTo Error_Handler

   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, Data() As String, j As Integer
   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer

   szSQL = "SELECT C.NominalCode AS BNC, N.Name AS BNN " & _
           "FROM tlbClientBanks AS C, NominalLedger AS N " & _
           "WHERE C.CLIENT_ID = N.ClientID AND C.NominalCode = N.Code AND C.CLIENT_ID <> '';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   iTotalRow = adoRST.RecordCount
   iTotalCol = adoRST.Fields.Count
   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String

   For i = 0 To iTotalRow
       For j = 0 To iTotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields.Item(j).Value), "", adoRST.Fields.Item(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cboBC.Column() = Data()

NoRes:
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub

Error_Handler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   Set adoRST = Nothing
End Sub

Private Sub fraButtons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
   
   If fraAddTrans.Visible Then fraAddTrans.Visible = False
End Sub

Private Sub lblBankRec_Click(Index As Integer)
    If Index >= 0 And Index <= 6 Then
'         Me.MousePointer = vbHourglass
'       If Index = 0 Then
'             SortingGrid flxBankPay, Index + 1, bSortingCol(Index), "Integer"
'      Else
'            SortingGrid flxBankPay, Index + 1, bSortingCol(Index)
'      End If
'
'      bSortingCol(Index) = IIf(bSortingCol(Index), False, True)
'
       ' LblSortingClicked Index, lblBankRec, 0, 6
'        Me.MousePointer = vbArrow
            lblBankRec(Index).FontBold = Not lblBankRec(Index).FontBold
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            ConfigFlxBankPay
            LoadFlxBankPay1 adoConn, Index, IIf(lblBankRec(Index).FontBold, "DESC", "ASC")
            adoConn.Close
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
Public Sub LoadFlxBankPay1(adoConn As ADODB.Connection, j As Integer, strSort As String)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRST As New ADODB.Recordset
   Dim strWhereBC As String
   Dim strWhereType As String
'  Column Heading: "Empty|<SL No|<Date|<Bank|<PropName|<NC|<Fund|<Ref|<Desc|>Amount|TranID|PD"
''                     ^      ^      ^     ^      ^      ^     ^    ^     ^       ^     ^
'   szSQL = "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
'                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, BP.MY_ID, C.CLIENT_ID, " & _
'                  "TT.DESCRIPTION AS Type1, BP.PROJ_REF AS Rfn, BP.PropertyID, BP.UNIT_ID, BP.NOMINAL_CODE, " & _
'                  "BP.BANK_AC AS ACC, C.Bank_AC_Name, BP.DEPT_ID, BP.NET_AMOUNT, BP.VAT, BP.ReconNow, " & _
'                  "BP.TAX_CODE, P.PropertyName, BP.PostingDate " & _
'           "FROM ((tlbBankPayment AS BP INNER JOIN tlbTransactionTypes AS TT ON BP.TransactionType = TT.TYPE_ID) INNER JOIN " & _
'                  "tlbClientBanks AS C ON BP.BANK_AC = C.NominalCode) LEFT JOIN " & _
'                  "Property AS P ON BP.PropertyID = P.PropertyID " & _
'           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
'                  "C.CLIENT_ID = BP.ClientID AND " & _
'                  "BP.UPDATE_SAGE = FALSE " & _
'           "ORDER BY TRANS, CLNG(TRAN_ID)"
'added client table on query by anol 05122016
',(Select UnitName FROM UNITS where UnitNumber=UNIT_ID) AS UNITNAME
    Dim szOrderby, strWhereCL, strWherePR, strWhereDt As String
    If j = 0 Then
        szOrderby = "ORDER BY TRANS, CLNG(TRAN_ID) " & strSort
    End If
    If j = 1 Then
        szOrderby = "ORDER BY TRANS,BP.TRAN_DATE " & strSort
    End If
    If j = 2 Then
        szOrderby = "ORDER BY TRANS,BP.BANK_AC " & strSort
    End If
    If j = 3 Then
        szOrderby = "ORDER BY TRANS,BP.PropertyID " & strSort
    End If
    If j = 4 Then
        szOrderby = "ORDER BY TRANS,BP.NOMINAL_CODE " & strSort
    End If
    If j = 5 Then
        szOrderby = "ORDER BY TRANS,BP.DEPT_ID " & strSort
    End If
    If j = 6 Then
        szOrderby = "ORDER BY TRANS,BP.PROJ_REF " & strSort
    End If
    
    If txtClientList.Tag <> "ALL" Then
        strWhereCL = " AND BP.ClientID='" & txtClientList.Tag & "' "
    End If
    If txtBankAccountFilter.text <> "ALL" Then
        strWhereBC = " AND BP.BANK_AC = '" & txtBankAccountFilter.text & "' "
    End If
    If txtTransTypeFilter.Tag <> "ALL" Then
        'BANK TRANSFER
        If txtTransTypeFilter.Tag = "BT" Then
             strWhereType = " AND BP.Description = 'BANK TRANSFER' "
        Else
             strWhereType = " AND BP.TRANS = '" & txtTransTypeFilter.Tag & "' "
        End If
    End If
    If txtPoperty.Tag <> "ALL" Then
        strWherePR = " AND BP.PropertyID='" & txtPoperty.Tag & "' "
    End If
    If Trim(txtDateFrom.text) <> "" And txtDateTo.text <> "" Then
    'Format(txtDateFrom.text, "dd mmmm yyyy")
        strWhereDt = " AND BP.TRAN_DATE>=#" & Format(txtDateFrom.text, "dd mmmm yyyy") & "# AND BP.TRAN_DATE<=#" & Format(txtDateTo.text, "dd mmmm yyyy") & "# "
    End If
            szSQL = "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, BP.MY_ID, C.CLIENT_ID, " & _
                  "TT.DESCRIPTION AS Type1, BP.PROJ_REF AS Rfn, BP.PropertyID, BP.UNIT_ID, BP.NOMINAL_CODE, " & _
                  "BP.BANK_AC AS ACC, NL.Name AS NN, BP.DEPT_ID, BP.NET_AMOUNT, BP.VAT, BP.ReconNow, " & _
                  "BP.TAX_CODE, P.PropertyName, BP.PostingDate, D.ClientName,(Select UnitName FROM UNITS where " & _
                  "UnitNumber=BP.UNIT_ID) as UnitName,(Select FundName from FUND where CSTR(FundID)=BP.DEPT_ID) AS FUNDNAME,(Select Name FROM NominalLedger where Code=BP.NOMINAL_CODE AND ClientID=BP.ClientID) as Name1 " & _
           "FROM ((((tlbBankPayment AS BP INNER JOIN CLIENT D ON D.ClientID=BP.ClientID) INNER JOIN tlbTransactionTypes AS TT ON BP.TransactionType = TT.TYPE_ID)  INNER JOIN " & _
                  "tlbClientBanks AS C ON BP.BANK_AC = C.NominalCode) INNER JOIN NominalLedger NL ON NL.ClientID=C.CLIENT_ID AND NL.CODE=C.NominalCode) LEFT JOIN " & _
                  "Property AS P ON BP.PropertyID = P.PropertyID " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                  "C.CLIENT_ID = BP.ClientID " & strWhereCL & strWherePR & strWhereDt & strWhereBC & strWhereType & " AND " & _
                  "BP.UPDATE_SAGE = FALSE " & _
           szOrderby

'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   flxBankPay.Clear
   flxBankPay.Rows = 2
   While Not adoRST.EOF
      flxBankPay.TextMatrix(i, 1) = adoRST.Fields.Item("Type2").Value & _
                                    adoRST.Fields.Item("T_ID").Value                            'SL No
      flxBankPay.TextMatrix(i, 2) = adoRST.Fields.Item("RDate").Value                           'Date
      flxBankPay.TextMatrix(i, 3) = adoRST.Fields.Item("ACC").Value                             'Account Name
      flxBankPay.TextMatrix(i, 4) = IIf(IsNull(adoRST!PropertyName), "", adoRST!PropertyName)   'Property Name
      flxBankPay.TextMatrix(i, 5) = IIf(IsNull(adoRST!Nominal_code), "", adoRST!Nominal_code)   'NOMINAL_CODE
      flxBankPay.TextMatrix(i, 6) = adoRST.Fields.Item("DEPT_ID").Value                         'Fund
      flxBankPay.TextMatrix(i, 7) = adoRST.Fields.Item("Rfn").Value                             'Ref
      flxBankPay.TextMatrix(i, 8) = adoRST.Fields.Item("Details").Value                         'Desc
      flxBankPay.TextMatrix(i, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")          'Amt
      flxBankPay.TextMatrix(i, 10) = adoRST.Fields.Item("MY_ID").Value                          'TranID
      flxBankPay.TextMatrix(i, 11) = IIf(IsNull(adoRST!propertyID), "", adoRST!propertyID)      'PropertyID
      flxBankPay.TextMatrix(i, 12) = IIf(IsNull(adoRST!UNIT_ID), "", adoRST!UNIT_ID)            'UNIT_ID
      flxBankPay.TextMatrix(i, 13) = IIf(IsNull(adoRST!CLIENT_ID), "", adoRST!CLIENT_ID)        'CLIENT ID
      flxBankPay.TextMatrix(i, 14) = IIf(IsNull(adoRST!ACC), "", adoRST!ACC)                    'BANK_AC
      flxBankPay.TextMatrix(i, 15) = adoRST.Fields.Item("NET_AMOUNT").Value                     'NET_AMOUNT
      flxBankPay.TextMatrix(i, 16) = adoRST.Fields.Item("VAT").Value                            'VAT
      flxBankPay.TextMatrix(i, 17) = IIf(IsNull(adoRST!ReconNow) Or adoRST!ReconNow = "", "N", "Y") 'IIf(IsNull(adoRst!ReconNow), "N", "Y")
      flxBankPay.TextMatrix(i, 18) = IIf(IsNull(adoRST!TAX_CODE), "", adoRST!TAX_CODE)
      flxBankPay.TextMatrix(i, 19) = IIf(IsNull(adoRST!postingDate), adoRST.Fields.Item("RDate").Value, adoRST!postingDate)    'Posting Date
      flxBankPay.TextMatrix(i, 20) = adoRST!ClientName    'CLIENT Name
      flxBankPay.TextMatrix(i, 21) = adoRST!NN    'Nominal Bank AC  Name
'      Dim rsUnit As New ADODB.Recordset
'      rsUnit.Open "Select UnitName FROM UNITS where UnitNumber='" & IIf(IsNull(adoRst!UNIT_ID), "", adoRst!UNIT_ID) & "'", adoConn, adOpenKeyset, adLockReadOnly
'      If Not rsUnit.EOF Then
'        flxBankPay.TextMatrix(i, 22) = "UNit Name" 'IIf(IsNull(rsUnit.Fields(0).Value), "", rsUnit.Fields(0).Value)
'      End If
        flxBankPay.TextMatrix(i, 22) = IIf(IsNull(adoRST!UnitName), "", adoRST!UnitName)
        flxBankPay.TextMatrix(i, 23) = IIf(IsNull(adoRST!FundName), "", adoRST!FundName)
        flxBankPay.TextMatrix(i, 24) = IIf(IsNull(adoRST!Name1), "", adoRST!Name1) 'Nominal account Name
'      rsUnit.Close
      adoRST.MoveNext
      If Not adoRST.EOF Then flxBankPay.AddItem ""
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing
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
    Dim adoConn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        adoConn.Open getConnectionString
        If Len(txtSearchNo.text) > 0 Then
            LoadFlxBankPay adoConn, "1"
        Else
            LoadFlxBankPay adoConn, ""
        End If
    '    fmeLoading.Visible = False
        adoConn.Close
        Set adoConn = Nothing
        If Len(txtSearchNo.text) > 0 Then
            cmdSearch.Caption = "Clear Sea&rch"
        Else
            cmdSearch.Caption = "Sea&rch"
        End If
    ElseIf tabPayment.Tab = 1 Then
        adoConn.Open getConnectionString
        If Len(txtSearchNo.text) > 0 Then
            LoadflxBankPayHist adoConn, "1"
        Else
            LoadflxBankPayHist adoConn, ""
        End If
        adoConn.Close
        Set adoConn = Nothing
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
    Dim adoConn As New ADODB.Connection
    If tabPayment.Tab = 0 Then
        adoConn.Open getConnectionString
        If Len(txtSearchRef.text) > 0 Then
            LoadFlxBankPay adoConn, "2"
        Else
            LoadFlxBankPay adoConn, ""
        End If
    '    fmeLoading.Visible = False
        adoConn.Close
        Set adoConn = Nothing
        If Len(txtSearchRef.text) > 0 Then
            cmdSearch.Caption = "Clear Sea&rch"
        Else
            cmdSearch.Caption = "Sea&rch"
        End If
    ElseIf tabPayment.Tab = 1 Then
        adoConn.Open getConnectionString
        If Len(txtSearchRef.text) > 0 Then
            LoadflxBankPayHist adoConn, "2"
        Else
            LoadflxBankPayHist adoConn, ""
        End If
        adoConn.Close
        Set adoConn = Nothing
        If Len(txtSearchRef.text) > 0 Then
            cmdSearchHistory.Caption = "Clear Sea&rch"
        Else
            cmdSearchHistory.Caption = "Sea&rch"
        End If
     End If
End Sub
