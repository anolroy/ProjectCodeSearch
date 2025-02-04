VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{EC4A06C3-9499-4BA0-8D3C-6EDE133B1673}#1.1#0"; "HoverButtons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDemands1 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demand & Receipt"
   ClientHeight    =   8205
   ClientLeft      =   540
   ClientTop       =   690
   ClientWidth     =   13440
   Icon            =   "Demand1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13440
   Begin VB.TextBox txtSelectedPrintRef 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   -1080
      Locked          =   -1  'True
      TabIndex        =   82
      Text            =   "txtSelectedPrintRef"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin TabDlg.SSTab tabDmdRcpt 
      Height          =   8055
      Left            =   75
      TabIndex        =   9
      Top             =   75
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   14208
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   16768960
      TabCaption(0)   =   "Demands"
      TabPicture(0)   =   "Demand1.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkSelectAllDemands"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flxDemands"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraMain"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDetails"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraPrintChoice"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraAutoDemandChoice"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraCreateManualDemand"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraEditDemandWindow"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraInvCrChoice"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Demand History"
      TabPicture(1)   =   "Demand1.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Receipts"
      TabPicture(2)   =   "Demand1.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabPayment"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab tabPayment 
         Height          =   7335
         Left            =   -74640
         TabIndex        =   136
         Top             =   480
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   12938
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Tenant Receipt"
         TabPicture(0)   =   "Demand1.frx":091E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dtSPDate"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame8"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "flxSPayment"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtSPayment"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "fraListNC"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Bank Receipt"
         TabPicture(1)   =   "Demand1.frx":093A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "flxBankPay(0)"
         Tab(1).Control(1)=   "fraBkInput(0)"
         Tab(1).Control(2)=   "txtBkTotalVat(0)"
         Tab(1).Control(3)=   "txtBkTotalNet(0)"
         Tab(1).Control(4)=   "fraListBk(0)"
         Tab(1).Control(5)=   "Frame5(2)"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Bank Payment"
         TabPicture(2)   =   "Demand1.frx":0956
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame5(1)"
         Tab(2).Control(1)=   "fraListBk(1)"
         Tab(2).Control(2)=   "txtBkTotalNet(1)"
         Tab(2).Control(3)=   "txtBkTotalVat(1)"
         Tab(2).Control(4)=   "fraBkInput(1)"
         Tab(2).Control(5)=   "flxBankPay(1)"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Bank Transfer"
         TabPicture(3)   =   "Demand1.frx":0972
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5(3)"
         Tab(3).Control(1)=   "MSHFlexGrid2"
         Tab(3).Control(2)=   "cboBkTrDept"
         Tab(3).Control(3)=   "txtBkTrAmt"
         Tab(3).Control(4)=   "txtBkTrDes"
         Tab(3).Control(5)=   "txtBkTrRef"
         Tab(3).Control(6)=   "cboBkTrAcTo"
         Tab(3).Control(7)=   "cboBkTrAcFr"
         Tab(3).Control(8)=   "txtBkTrDate"
         Tab(3).Control(9)=   "Label17"
         Tab(3).Control(10)=   "Label16"
         Tab(3).Control(11)=   "Label15"
         Tab(3).Control(12)=   "Label14"
         Tab(3).Control(13)=   "Label13"
         Tab(3).Control(14)=   "Label12"
         Tab(3).Control(15)=   "Label11"
         Tab(3).ControlCount=   16
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   680
            Index           =   3
            Left            =   -66000
            TabIndex        =   284
            Top             =   6480
            Width           =   3855
            Begin VB.CommandButton Command10 
               BackColor       =   &H00C0E5BA&
               Caption         =   "&Save"
               Height          =   400
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   286
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton Command9 
               BackColor       =   &H00C0E5BA&
               Caption         =   "Cancel"
               Height          =   400
               Left            =   2160
               Style           =   1  'Graphical
               TabIndex        =   285
               Top             =   120
               Visible         =   0   'False
               Width           =   1575
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3495
            Left            =   -74640
            TabIndex        =   283
            Top             =   2760
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   6165
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame fraListNC 
            BackColor       =   &H00FDEDED&
            BorderStyle     =   0  'None
            Height          =   2655
            Left            =   2880
            TabIndex        =   264
            Top             =   2160
            Visible         =   0   'False
            Width           =   4815
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxListNC 
               Height          =   1935
               Left            =   75
               TabIndex        =   265
               Top             =   300
               Width           =   4650
               _ExtentX        =   8202
               _ExtentY        =   3413
               _Version        =   393216
               ForeColor       =   0
               FixedCols       =   0
               BackColorSel    =   7573887
               GridColor       =   -2147483635
               SelectionMode   =   1
               Appearance      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin HoverButton.HoverControl cmdOKFlxNC 
               Height          =   300
               Left            =   3840
               TabIndex        =   266
               Top             =   2300
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               BackColor       =   -2147483624
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "&OK"
            End
            Begin VB.Label lblPayeeFlxConfigured 
               Caption         =   "NOT"
               Height          =   495
               Index           =   1
               Left            =   1200
               TabIndex        =   268
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label lblFlxPayee 
               Caption         =   "EMPTY"
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   267
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Image imgFramListCoseNC 
               Height          =   240
               Left            =   4515
               Stretch         =   -1  'True
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00B5BAA5&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   1
            Left            =   -74400
            TabIndex        =   256
            Top             =   6480
            Width           =   12375
            Begin VB.CommandButton cmdEditBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Edit"
               Height          =   400
               Index           =   1
               Left            =   1920
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   258
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdNewBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&New Pyament"
               Height          =   400
               Index           =   1
               Left            =   120
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   257
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdCloseBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "C&lose"
               Height          =   400
               Index           =   1
               Left            =   10680
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   261
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdSaveBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               Height          =   400
               Index           =   1
               Left            =   7200
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   260
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdCancelBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancel"
               Height          =   400
               Index           =   1
               Left            =   3720
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   259
               Top             =   120
               Width           =   1450
            End
         End
         Begin VB.Frame fraListBk 
            BackColor       =   &H00FDEDED&
            BorderStyle     =   0  'None
            Height          =   2655
            Index           =   1
            Left            =   -72360
            TabIndex        =   251
            Top             =   2760
            Visible         =   0   'False
            Width           =   4815
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxListBk 
               Height          =   1935
               Index           =   1
               Left            =   75
               TabIndex        =   252
               Top             =   300
               Width           =   4650
               _ExtentX        =   8202
               _ExtentY        =   3413
               _Version        =   393216
               ForeColor       =   0
               FixedCols       =   0
               BackColorSel    =   7573887
               GridColor       =   -2147483635
               SelectionMode   =   1
               Appearance      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin HoverButton.HoverControl cmdOKFlxBk 
               Height          =   300
               Index           =   1
               Left            =   3840
               TabIndex        =   253
               Top             =   2300
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "&OK"
            End
            Begin VB.Label lblPayeeFlxConfigured 
               Caption         =   "NOT"
               Height          =   495
               Index           =   0
               Left            =   1200
               TabIndex        =   255
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label lblFlxPayee 
               Caption         =   "EMPTY"
               Height          =   255
               Index           =   0
               Left            =   2280
               TabIndex        =   254
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Image imgFramListCoseBk 
               Height          =   240
               Index           =   1
               Left            =   4515
               Stretch         =   -1  'True
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.TextBox txtBkTotalNet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   -64800
            TabIndex        =   250
            Text            =   "0.00"
            Top             =   6000
            Width           =   975
         End
         Begin VB.TextBox txtBkTotalVat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   -63120
            TabIndex        =   249
            Text            =   "0.00"
            Top             =   6000
            Width           =   975
         End
         Begin VB.Frame fraBkInput 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   1
            Left            =   -74400
            TabIndex        =   212
            Top             =   600
            Width           =   12375
            Begin VB.TextBox txtBkAc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   227
               Top             =   120
               Width           =   2000
            End
            Begin VB.TextBox txtDateBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   226
               Top             =   420
               Width           =   2000
            End
            Begin VB.TextBox txtTypeBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   225
               Top             =   720
               Width           =   2000
            End
            Begin VB.ListBox lstTypeBk 
               Height          =   450
               Index           =   1
               Left            =   -2040
               TabIndex        =   224
               Top             =   1200
               Visible         =   0   'False
               Width           =   2240
            End
            Begin VB.TextBox txtUnitBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   223
               Top             =   1020
               Width           =   2000
            End
            Begin VB.TextBox txtDeptBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   222
               Top             =   420
               Width           =   2000
            End
            Begin VB.TextBox txtNCBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   221
               Top             =   120
               Width           =   2000
            End
            Begin VB.TextBox txtProjBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   220
               Top             =   720
               Width           =   2000
            End
            Begin VB.TextBox txtCCBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   219
               Top             =   1020
               Width           =   2000
            End
            Begin VB.TextBox txtNetBk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   9000
               TabIndex        =   218
               Top             =   420
               Width           =   1300
            End
            Begin VB.TextBox txtDetailsBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   9000
               TabIndex        =   217
               Top             =   120
               Width           =   3195
            End
            Begin VB.TextBox txtVatBk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   9000
               TabIndex        =   216
               Top             =   720
               Width           =   1300
            End
            Begin VB.TextBox txtRecharge 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   9000
               TabIndex        =   215
               Text            =   "NO"
               Top             =   1020
               Width           =   435
            End
            Begin VB.ListBox lstYNBk 
               Height          =   450
               Index           =   1
               Left            =   9960
               TabIndex        =   214
               Top             =   840
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CommandButton cmdUpdateBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Update Record"
               Height          =   375
               Index           =   1
               Left            =   10920
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   213
               Top             =   960
               Width           =   1335
            End
            Begin HoverButton.HoverControl cmdBkList 
               Height          =   285
               Index           =   1
               Left            =   2820
               TabIndex        =   228
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdTransListBk 
               Height          =   285
               Index           =   1
               Left            =   2820
               TabIndex        =   229
               Top             =   720
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdUnitListBk 
               Height          =   285
               Index           =   1
               Left            =   2820
               TabIndex        =   230
               Top             =   1020
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdNCBk 
               Height          =   285
               Index           =   1
               Left            =   7260
               TabIndex        =   231
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdProjBk 
               Height          =   285
               Index           =   1
               Left            =   7260
               TabIndex        =   232
               Top             =   720
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdCCBk 
               Height          =   285
               Index           =   1
               Left            =   7260
               TabIndex        =   233
               Top             =   1020
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdTaxListBk 
               Height          =   255
               Index           =   1
               Left            =   10320
               TabIndex        =   234
               Top             =   720
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "..."
            End
            Begin HoverButton.HoverControl cmdDeptBk 
               Height          =   285
               Index           =   1
               Left            =   7260
               TabIndex        =   235
               Top             =   420
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdRechargeBk 
               Height          =   285
               Index           =   1
               Left            =   9440
               TabIndex        =   236
               Top             =   1020
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BankA/C:"
               Height          =   195
               Index           =   24
               Left            =   120
               TabIndex        =   248
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date:"
               Height          =   195
               Index           =   23
               Left            =   120
               TabIndex        =   247
               Top             =   420
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type:"
               Height          =   195
               Index           =   22
               Left            =   120
               TabIndex        =   246
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit:"
               Height          =   195
               Index           =   21
               Left            =   120
               TabIndex        =   245
               Top             =   1020
               Width           =   330
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dept:"
               Height          =   195
               Index           =   20
               Left            =   4560
               TabIndex        =   244
               Top             =   420
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N/C:"
               Height          =   195
               Index           =   19
               Left            =   4560
               TabIndex        =   243
               Top             =   120
               Width           =   345
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Proj.:"
               Height          =   195
               Index           =   18
               Left            =   4560
               TabIndex        =   242
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CC:"
               Height          =   195
               Index           =   17
               Left            =   4560
               TabIndex        =   241
               Top             =   1020
               Width           =   255
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Net:"
               Height          =   195
               Index           =   16
               Left            =   8160
               TabIndex        =   240
               Top             =   420
               Width           =   300
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Details:"
               Height          =   195
               Index           =   15
               Left            =   8160
               TabIndex        =   239
               Top             =   120
               Width           =   525
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VAT:"
               Height          =   195
               Index           =   14
               Left            =   8160
               TabIndex        =   238
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recharge:"
               Height          =   195
               Index           =   1
               Left            =   8160
               TabIndex        =   237
               Top             =   1020
               Width           =   750
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00B5BAA5&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   2
            Left            =   -74400
            TabIndex        =   205
            Top             =   6480
            Width           =   12375
            Begin VB.CommandButton cmdCancelBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancel"
               Height          =   400
               Index           =   0
               Left            =   3720
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   208
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdSaveBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               Height          =   400
               Index           =   0
               Left            =   6960
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   209
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdCloseBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "C&lose"
               Height          =   400
               Index           =   0
               Left            =   10680
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   210
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdNewBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&New Pyament"
               Height          =   400
               Index           =   0
               Left            =   120
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   206
               Top             =   120
               Width           =   1450
            End
            Begin VB.CommandButton cmdEditBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Edit"
               Height          =   400
               Index           =   0
               Left            =   1920
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   207
               Top             =   120
               Width           =   1450
            End
         End
         Begin VB.Frame fraListBk 
            BackColor       =   &H00FDEDED&
            BorderStyle     =   0  'None
            Height          =   2655
            Index           =   0
            Left            =   -72360
            TabIndex        =   200
            Top             =   2760
            Visible         =   0   'False
            Width           =   4815
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxListBk 
               Height          =   1935
               Index           =   0
               Left            =   75
               TabIndex        =   201
               Top             =   300
               Width           =   4650
               _ExtentX        =   8202
               _ExtentY        =   3413
               _Version        =   393216
               ForeColor       =   0
               FixedCols       =   0
               BackColorSel    =   7573887
               GridColor       =   -2147483635
               SelectionMode   =   1
               Appearance      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin HoverButton.HoverControl cmdOKFlxBk 
               Height          =   300
               Index           =   0
               Left            =   3840
               TabIndex        =   202
               Top             =   2300
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "&OK"
            End
            Begin VB.Image imgFramListCoseBk 
               Height          =   240
               Index           =   0
               Left            =   4515
               Stretch         =   -1  'True
               Top             =   0
               Width           =   240
            End
            Begin VB.Label lblFlxPayee 
               Caption         =   "EMPTY"
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   204
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label lblPayeeFlxConfigured 
               Caption         =   "NOT"
               Height          =   495
               Index           =   2
               Left            =   1200
               TabIndex        =   203
               Top             =   1680
               Width           =   1095
            End
         End
         Begin VB.TextBox txtBkTotalNet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   -64800
            TabIndex        =   199
            Text            =   "0.00"
            Top             =   6000
            Width           =   975
         End
         Begin VB.TextBox txtBkTotalVat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   -63120
            TabIndex        =   198
            Text            =   "0.00"
            Top             =   6000
            Width           =   975
         End
         Begin VB.Frame fraBkInput 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   0
            Left            =   -74400
            TabIndex        =   161
            Top             =   600
            Width           =   12375
            Begin VB.CommandButton cmdUpdateBk 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Update Record"
               Height          =   375
               Index           =   0
               Left            =   10920
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   960
               Width           =   1335
            End
            Begin VB.ListBox lstYNBk 
               Height          =   450
               Index           =   0
               Left            =   9960
               TabIndex        =   175
               Top             =   840
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtRecharge 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   9000
               TabIndex        =   174
               Text            =   "NO"
               Top             =   1020
               Width           =   435
            End
            Begin VB.TextBox txtVatBk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   9000
               TabIndex        =   173
               Top             =   720
               Width           =   1300
            End
            Begin VB.TextBox txtDetailsBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   9000
               TabIndex        =   172
               Top             =   120
               Width           =   3195
            End
            Begin VB.TextBox txtNetBk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   9000
               TabIndex        =   171
               Top             =   420
               Width           =   1300
            End
            Begin VB.TextBox txtCCBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   5280
               TabIndex        =   170
               Top             =   1020
               Width           =   2000
            End
            Begin VB.TextBox txtProjBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   5280
               TabIndex        =   169
               Top             =   720
               Width           =   2000
            End
            Begin VB.TextBox txtNCBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   5280
               TabIndex        =   168
               Top             =   120
               Width           =   2000
            End
            Begin VB.TextBox txtDeptBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   5280
               TabIndex        =   167
               Top             =   420
               Width           =   2000
            End
            Begin VB.TextBox txtUnitBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   840
               TabIndex        =   166
               Top             =   1020
               Width           =   2000
            End
            Begin VB.ListBox lstTypeBk 
               Height          =   450
               Index           =   0
               Left            =   -2040
               TabIndex        =   165
               Top             =   1200
               Visible         =   0   'False
               Width           =   2240
            End
            Begin VB.TextBox txtTypeBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   840
               TabIndex        =   164
               Top             =   720
               Width           =   2000
            End
            Begin VB.TextBox txtDateBk 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   840
               TabIndex        =   163
               Top             =   420
               Width           =   2000
            End
            Begin VB.TextBox txtBkAc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   840
               TabIndex        =   162
               Top             =   120
               Width           =   2000
            End
            Begin HoverButton.HoverControl cmdBkList 
               Height          =   285
               Index           =   0
               Left            =   2820
               TabIndex        =   177
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdTransListBk 
               Height          =   285
               Index           =   0
               Left            =   2820
               TabIndex        =   178
               Top             =   720
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdUnitListBk 
               Height          =   285
               Index           =   0
               Left            =   2820
               TabIndex        =   179
               Top             =   1020
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdNCBk 
               Height          =   285
               Index           =   0
               Left            =   7260
               TabIndex        =   180
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdProjBk 
               Height          =   285
               Index           =   0
               Left            =   7260
               TabIndex        =   181
               Top             =   720
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdCCBk 
               Height          =   285
               Index           =   0
               Left            =   7260
               TabIndex        =   182
               Top             =   1020
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdTaxListBk 
               Height          =   255
               Index           =   0
               Left            =   10320
               TabIndex        =   183
               Top             =   720
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "..."
            End
            Begin HoverButton.HoverControl cmdDeptBk 
               Height          =   285
               Index           =   0
               Left            =   7260
               TabIndex        =   184
               Top             =   420
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin HoverButton.HoverControl cmdRechargeBk 
               Height          =   285
               Index           =   0
               Left            =   9440
               TabIndex        =   185
               Top             =   1020
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recharge:"
               Height          =   195
               Index           =   13
               Left            =   8160
               TabIndex        =   197
               Top             =   1020
               Width           =   750
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VAT:"
               Height          =   195
               Index           =   12
               Left            =   8160
               TabIndex        =   196
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Details:"
               Height          =   195
               Index           =   11
               Left            =   8160
               TabIndex        =   195
               Top             =   120
               Width           =   525
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Net:"
               Height          =   195
               Index           =   10
               Left            =   8160
               TabIndex        =   194
               Top             =   420
               Width           =   300
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CC:"
               Height          =   195
               Index           =   9
               Left            =   4560
               TabIndex        =   193
               Top             =   1020
               Width           =   255
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Proj.:"
               Height          =   195
               Index           =   8
               Left            =   4560
               TabIndex        =   192
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N/C:"
               Height          =   195
               Index           =   7
               Left            =   4560
               TabIndex        =   191
               Top             =   120
               Width           =   345
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dept:"
               Height          =   195
               Index           =   6
               Left            =   4560
               TabIndex        =   190
               Top             =   420
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit:"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   189
               Top             =   1020
               Width           =   330
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type:"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   188
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date:"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   187
               Top             =   420
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BankA/C:"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   186
               Top             =   120
               Width           =   705
            End
         End
         Begin VB.TextBox txtSPayment 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10920
            TabIndex        =   142
            Top             =   1920
            Visible         =   0   'False
            Width           =   1335
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
            Height          =   4335
            Left            =   360
            TabIndex        =   160
            Top             =   1680
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   7646
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            GridLinesFixed  =   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   240
            TabIndex        =   149
            Top             =   480
            Width           =   12975
            Begin HoverButton.HoverControl cmdNC 
               Height          =   285
               Left            =   8060
               TabIndex        =   141
               Top             =   600
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               BackColor       =   -2147483624
               ForeColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "V"
            End
            Begin VB.TextBox txtNominalCodeTR 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   285
               Left            =   6240
               TabIndex        =   140
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtChqNo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   10320
               TabIndex        =   150
               Top             =   120
               Width           =   2175
            End
            Begin VB.TextBox txtSPaymentTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   10320
               Locked          =   -1  'True
               TabIndex        =   151
               Text            =   "0.00"
               Top             =   600
               Width           =   2175
            End
            Begin VB.TextBox txtSPDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   285
               Left            =   6240
               TabIndex        =   139
               Top             =   120
               Width           =   1815
            End
            Begin VB.ComboBox cmbTenant 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   138
               Top             =   600
               Width           =   3495
            End
            Begin VB.ComboBox cmbBankAc 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   137
               Top             =   120
               Width           =   3495
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nominal Code"
               Height          =   195
               Left            =   5040
               TabIndex        =   263
               Top             =   600
               Width           =   990
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Receipt Date"
               Height          =   195
               Left            =   5040
               TabIndex        =   158
               Top             =   120
               Width           =   945
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank Balance               £"
               Height          =   195
               Left            =   8520
               TabIndex        =   157
               Top             =   165
               Width           =   1770
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Receipt Amount   £"
               Height          =   195
               Left            =   8520
               TabIndex        =   155
               Top             =   600
               Width           =   1770
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tenant"
               Height          =   195
               Left            =   240
               TabIndex        =   153
               Top             =   600
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank A/C"
               Height          =   195
               Index           =   6
               Left            =   240
               TabIndex        =   152
               Top             =   120
               Width           =   705
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   0
            Left            =   240
            TabIndex        =   147
            Top             =   6600
            Width           =   12975
            Begin VB.CommandButton Command6 
               BackColor       =   &H00C0E5BA&
               Caption         =   "Allocate"
               Height          =   400
               Left            =   7320
               Style           =   1  'Graphical
               TabIndex        =   156
               Top             =   120
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00C0E5BA&
               Caption         =   "Receipt on A/C"
               Height          =   400
               Left            =   5640
               Style           =   1  'Graphical
               TabIndex        =   154
               Top             =   120
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.CommandButton cmdSPDiscard 
               BackColor       =   &H00C0E5BA&
               Caption         =   "&Discard"
               Height          =   400
               Left            =   9600
               Style           =   1  'Graphical
               TabIndex        =   148
               Top             =   120
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.CommandButton cmdSPSave 
               BackColor       =   &H00C0E5BA&
               Caption         =   "&Save"
               Height          =   400
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   143
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton cmdSPayAll 
               BackColor       =   &H00C0E5BA&
               Caption         =   "Pay &All"
               Height          =   400
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   145
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton cmdSPClose 
               BackColor       =   &H00C0E5BA&
               Caption         =   "C&lose"
               Height          =   400
               Left            =   11280
               Style           =   1  'Graphical
               TabIndex        =   146
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton cmdSPFull 
               BackColor       =   &H00C0E5BA&
               Caption         =   "Pay in &Full"
               Height          =   400
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   144
               Top             =   120
               Width           =   1575
            End
         End
         Begin MSComCtl2.MonthView dtSPDate 
            Height          =   2370
            Left            =   6480
            TabIndex        =   159
            Top             =   840
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483629
            Appearance      =   1
            StartOfWeek     =   20709378
            CurrentDate     =   38628
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankPay 
            Height          =   3795
            Index           =   0
            Left            =   -74400
            TabIndex        =   211
            Top             =   2160
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   6694
            _Version        =   393216
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankPay 
            Height          =   3795
            Index           =   1
            Left            =   -74400
            TabIndex        =   262
            Top             =   2160
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   6694
            _Version        =   393216
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSForms.ComboBox cboBkTrDept 
            Height          =   375
            Left            =   -66285
            TabIndex        =   282
            Top             =   2160
            Width           =   4215
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "7435;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBkTrAmt 
            Height          =   375
            Left            =   -66285
            TabIndex        =   281
            Top             =   1680
            Width           =   2055
            VariousPropertyBits=   746604571
            Size            =   "3625;661"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtBkTrDes 
            Height          =   975
            Left            =   -66285
            TabIndex        =   280
            Top             =   600
            Width           =   4215
            VariousPropertyBits=   -1400879077
            Size            =   "7435;1720"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBkTrRef 
            Height          =   375
            Left            =   -72960
            TabIndex        =   279
            Top             =   2160
            Width           =   4215
            VariousPropertyBits=   746604571
            Size            =   "7435;661"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboBkTrAcTo 
            Height          =   375
            Left            =   -72960
            TabIndex        =   278
            Top             =   1680
            Width           =   4215
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "7435;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboBkTrAcFr 
            Height          =   375
            Left            =   -72960
            TabIndex        =   277
            Top             =   1080
            Width           =   4215
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "7435;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBkTrDate 
            Height          =   375
            Left            =   -72960
            TabIndex        =   276
            Top             =   600
            Width           =   1935
            VariousPropertyBits=   746604571
            Size            =   "3413;661"
            SpecialEffect   =   6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label17 
            Height          =   195
            Left            =   -67680
            TabIndex        =   275
            Top             =   2160
            Width           =   885
            BackColor       =   16768960
            VariousPropertyBits=   276824083
            Caption         =   "Department:"
            Size            =   "1561;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label16 
            Height          =   195
            Left            =   -67680
            TabIndex        =   274
            Top             =   1680
            Width           =   1125
            BackColor       =   16768960
            VariousPropertyBits=   276824083
            Caption         =   "Payment Value:"
            Size            =   "1984;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label15 
            Height          =   195
            Left            =   -67680
            TabIndex        =   273
            Top             =   600
            Width           =   855
            BackColor       =   16768960
            VariousPropertyBits=   276824083
            Caption         =   "Description:"
            Size            =   "1508;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label14 
            Height          =   195
            Left            =   -74445
            TabIndex        =   272
            Top             =   2160
            Width           =   810
            BackColor       =   16768960
            VariousPropertyBits=   276824083
            Caption         =   "Reference:"
            Size            =   "1429;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label13 
            Height          =   195
            Left            =   -74445
            TabIndex        =   271
            Top             =   600
            Width           =   405
            BackColor       =   16768960
            VariousPropertyBits=   276824083
            Caption         =   "Date:"
            Size            =   "714;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label12 
            Height          =   195
            Left            =   -74445
            TabIndex        =   270
            Top             =   1680
            Width           =   900
            BackColor       =   16768960
            VariousPropertyBits=   276824083
            Caption         =   "Account To:"
            Size            =   "1587;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label11 
            Height          =   195
            Left            =   -74445
            TabIndex        =   269
            Top             =   1080
            Width           =   1050
            BackColor       =   16768960
            VariousPropertyBits=   276824083
            Caption         =   "Account From:"
            Size            =   "1852;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fraInvCrChoice 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Caption         =   "Automatic Demand Generate:"
         ForeColor       =   &H00FF00FF&
         Height          =   2055
         Left            =   5760
         TabIndex        =   124
         Top             =   -1560
         Visible         =   0   'False
         Width           =   2295
         Begin VB.OptionButton optManualAdjInv 
            BackColor       =   &H80000018&
            Caption         =   "Adjustment Invoice"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   120
            TabIndex        =   291
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optManualAdjCrNote 
            BackColor       =   &H80000018&
            Caption         =   "Adjustment Credit Note"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   120
            TabIndex        =   290
            Top             =   1320
            Width           =   1935
         End
         Begin VB.OptionButton optManualCrNote 
            BackColor       =   &H80000018&
            Caption         =   "Credit Note"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optManualInv 
            BackColor       =   &H80000018&
            Caption         =   "Invoice"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00CACAAA&
            Caption         =   "Different Due Date:"
            ForeColor       =   &H00000080&
            Height          =   855
            Left            =   -240
            TabIndex        =   125
            Top             =   2280
            Visible         =   0   'False
            Width           =   3135
            Begin VB.OptionButton Option4 
               BackColor       =   &H00CACAAA&
               Caption         =   "&Yes"
               Height          =   255
               Left            =   2280
               TabIndex        =   127
               Top             =   140
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00CACAAA&
               Caption         =   "&No"
               Height          =   255
               Left            =   2280
               TabIndex        =   126
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Consolidate different due    date st.(s) in single Demand?"
               Height          =   435
               Index           =   16
               Left            =   -120
               TabIndex        =   128
               Top             =   240
               Width           =   2235
            End
         End
         Begin HoverButton.HoverControl cmdManualDmdOk 
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   1680
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   -2147483647
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&OK"
         End
         Begin HoverButton.HoverControl cmdManualDmdCancel 
            Height          =   255
            Left            =   1320
            TabIndex        =   132
            Top             =   1680
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   -2147483647
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Close"
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "->>Transaction Type:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   133
            Top             =   0
            Width           =   1830
         End
      End
      Begin VB.Frame fraEditDemandWindow 
         BackColor       =   &H00ECECEC&
         Caption         =   "Edit Demand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   7575
         Left            =   11400
         TabIndex        =   96
         Top             =   2760
         Visible         =   0   'False
         Width           =   12375
         Begin VB.TextBox txtEditBatch 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5535
            Locked          =   -1  'True
            TabIndex        =   122
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtEditTenantName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   116
            Top             =   240
            Width           =   2655
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            TabIndex        =   109
            Top             =   6120
            Width           =   12135
            Begin VB.TextBox txtEditSubAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FCF4F5&
               Height          =   285
               Left            =   8400
               Locked          =   -1  'True
               TabIndex        =   113
               Top             =   120
               Width           =   900
            End
            Begin VB.TextBox txtEditSubVat 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FCF4F5&
               Height          =   285
               Left            =   9315
               Locked          =   -1  'True
               TabIndex        =   112
               Top             =   120
               Width           =   800
            End
            Begin VB.TextBox txtEditSubTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FCF4F5&
               Height          =   285
               Left            =   10125
               Locked          =   -1  'True
               TabIndex        =   111
               Top             =   120
               Width           =   900
            End
            Begin VB.TextBox txtEditAddNewSageText 
               Appearance      =   0  'Flat
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   900
               Locked          =   -1  'True
               TabIndex        =   110
               Top             =   120
               Width           =   2055
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Total:"
               Height          =   195
               Index           =   10
               Left            =   7440
               TabIndex        =   115
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sage Text:"
               Height          =   195
               Index           =   9
               Left            =   60
               TabIndex        =   114
               Top             =   120
               Width           =   780
            End
         End
         Begin MSComCtl2.MonthView dtEditDate 
            Height          =   2370
            Left            =   2280
            TabIndex        =   107
            Top             =   1320
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   16768960
            Appearance      =   1
            StartOfWeek     =   20709378
            CurrentDate     =   38651
         End
         Begin VB.TextBox txtEditUnit 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtEditAmount 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   7200
            TabIndex        =   105
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboEditType 
            Height          =   315
            Left            =   480
            TabIndex        =   104
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00FEE8E8&
            BorderStyle     =   0  'None
            Caption         =   "New Demand"
            Height          =   615
            Left            =   8760
            TabIndex        =   101
            Top             =   6720
            Width           =   3375
            Begin VB.CommandButton cmdCancelUpdate 
               Caption         =   "&Cancel"
               Height          =   375
               Left            =   1800
               TabIndex        =   103
               Top             =   120
               Width           =   1335
            End
            Begin VB.CommandButton cmdUpdateDemand 
               Caption         =   "&Update"
               Height          =   375
               Left            =   240
               TabIndex        =   102
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.TextBox txtEditDemandID 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   100
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtEditIssueDate 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtEditDescription 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   480
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   98
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cboEditVatCode 
            Height          =   315
            Left            =   8520
            TabIndex        =   97
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEditDemand 
            Height          =   5055
            Left            =   120
            TabIndex        =   108
            Top             =   960
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   8916
            _Version        =   393216
            BackColorBkg    =   16248815
            GridColor       =   -2147483635
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.MonthView dtEditIssueDate 
            Height          =   2370
            Left            =   9480
            TabIndex        =   117
            Top             =   600
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   16768960
            Appearance      =   1
            StartOfWeek     =   20709378
            CurrentDate     =   38655
         End
         Begin VB.Label lblTransactionType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblTransactionType"
            Height          =   195
            Left            =   4560
            TabIndex        =   292
            Top             =   600
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Batch:"
            Height          =   195
            Index           =   11
            Left            =   4560
            TabIndex        =   123
            Top             =   240
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Demand ID:"
            Height          =   195
            Index           =   15
            Left            =   135
            TabIndex        =   121
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit No.:"
            Height          =   195
            Index           =   14
            Left            =   8640
            TabIndex        =   120
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Date:"
            Height          =   195
            Index           =   12
            Left            =   8565
            TabIndex        =   118
            Top             =   600
            Width           =   810
         End
      End
      Begin VB.Frame fraCreateManualDemand 
         BackColor       =   &H00F0CFDF&
         Caption         =   "Generate Manual Demands"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   7575
         Left            =   1560
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   12135
         Begin VB.TextBox txtDate 
            BorderStyle     =   0  'None
            Height          =   260
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   289
            Top             =   4320
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtBatchID 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   288
            Text            =   "Batch"
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox cboVatCode 
            Height          =   315
            Left            =   7680
            TabIndex        =   83
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtAddNewDescription 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   1320
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtIssueDate 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   9360
            TabIndex        =   3
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtDemandID 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   600
            Width           =   2655
         End
         Begin VB.Frame fraAddNew 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "New Demand"
            Height          =   735
            Left            =   7320
            TabIndex        =   36
            Top             =   6720
            Width           =   4695
            Begin VB.CommandButton cmdAddNewDemand 
               Caption         =   "Add New"
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   200
               Width           =   1335
            End
            Begin VB.CommandButton cmdSaveNew 
               Caption         =   "&Save"
               Height          =   375
               Left            =   1680
               TabIndex        =   7
               Top             =   200
               Width           =   1335
            End
            Begin VB.CommandButton cmdCancelNew 
               Caption         =   "&Cancel"
               Height          =   375
               Left            =   3240
               TabIndex        =   8
               Top             =   200
               Width           =   1335
            End
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            Left            =   360
            TabIndex        =   4
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtAddNewAmount 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   6360
            TabIndex        =   35
            Top             =   1320
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtUnit 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   2655
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAddNewDemands 
            Height          =   5295
            Left            =   75
            TabIndex        =   37
            Top             =   960
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   9340
            _Version        =   393216
            GridColor       =   -2147483635
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   67
            Top             =   6120
            Width           =   11895
            Begin VB.TextBox txtAddNewSageText 
               Appearance      =   0  'Flat
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   900
               Locked          =   -1  'True
               TabIndex        =   74
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox txtSubTTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FCF4F5&
               Height          =   285
               Left            =   10125
               Locked          =   -1  'True
               TabIndex        =   70
               Top             =   240
               Width           =   900
            End
            Begin VB.TextBox txtSubTVAT 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FCF4F5&
               Height          =   285
               Left            =   9315
               Locked          =   -1  'True
               TabIndex        =   69
               Top             =   240
               Width           =   800
            End
            Begin VB.TextBox txtSubTAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FCF4F5&
               Height          =   285
               Left            =   8400
               Locked          =   -1  'True
               TabIndex        =   68
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sage Text:"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   75
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Total:"
               Height          =   195
               Index           =   3
               Left            =   7440
               TabIndex        =   71
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.ComboBox cboTenant 
            Height          =   315
            Left            =   1080
            TabIndex        =   0
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtTenantName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblDemandType 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   5760
            TabIndex        =   135
            Top             =   720
            Width           =   1740
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Demand Type:"
            Height          =   195
            Index           =   17
            Left            =   4560
            TabIndex        =   134
            Top             =   720
            Width           =   1050
         End
         Begin MSForms.CheckBox chkOldTenats 
            Height          =   255
            Left            =   3840
            TabIndex        =   95
            Top             =   240
            Width           =   2055
            BackColor       =   15781855
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3625;450"
            Value           =   "0"
            Caption         =   "Include Old Tenants"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Date:"
            Height          =   195
            Index           =   7
            Left            =   8445
            TabIndex        =   73
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit No.:"
            Height          =   195
            Index           =   2
            Left            =   8520
            TabIndex        =   39
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Demand ID:"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   38
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Details"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   91
         Top             =   4320
         Width           =   13575
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChildDemandHistory 
            Height          =   2655
            Left            =   75
            TabIndex        =   92
            Top             =   240
            Width           =   13335
            _ExtentX        =   23521
            _ExtentY        =   4683
            _Version        =   393216
            BackColor       =   15988477
            BackColorFixed  =   13883619
            BackColorBkg    =   13883619
            GridColor       =   -2147483635
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   615
         Left            =   -65400
         TabIndex        =   86
         Top             =   7320
         Width           =   4095
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FAD5DF&
            Caption         =   "Print Selected"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FAD5DF&
            Caption         =   "Print All"
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FAD5DF&
            Caption         =   "&All Unprinted"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   -74880
         TabIndex        =   84
         Top             =   480
         Width           =   13575
         Begin VB.TextBox txtSearchDemandID 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   80
            TabIndex        =   85
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandHistory 
            Height          =   3375
            Left            =   75
            TabIndex        =   90
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   5953
            _Version        =   393216
            GridColor       =   -2147483635
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000003&
            Caption         =   "*Type Demand ID here for search."
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Left            =   1920
            TabIndex        =   93
            Top             =   240
            Visible         =   0   'False
            Width           =   2430
         End
      End
      Begin VB.Frame fraAutoDemandChoice 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Caption         =   "Automatic Demand Generate:"
         ForeColor       =   &H00FF00FF&
         Height          =   1695
         Left            =   5520
         TabIndex        =   29
         Top             =   -1320
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Frame Frame4 
            BackColor       =   &H00CACAAA&
            Caption         =   "Different Due Date:"
            ForeColor       =   &H00000080&
            Height          =   855
            Left            =   120
            TabIndex        =   78
            Top             =   1800
            Visible         =   0   'False
            Width           =   3135
            Begin VB.OptionButton Option2 
               BackColor       =   &H00CACAAA&
               Caption         =   "&No"
               Height          =   255
               Left            =   2280
               TabIndex        =   80
               Top             =   480
               Width           =   615
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00CACAAA&
               Caption         =   "&Yes"
               Height          =   255
               Left            =   2280
               TabIndex        =   79
               Top             =   140
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Consolidate different due    date st.(s) in single Demand?"
               Height          =   435
               Index           =   8
               Left            =   -120
               TabIndex        =   81
               Top             =   240
               Width           =   2235
            End
         End
         Begin HoverButton.HoverControl cmdAutoDmdGenOk 
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1320
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   -2147483647
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&OK"
         End
         Begin VB.OptionButton optAutoGenConsolidated 
            BackColor       =   &H80000018&
            Caption         =   "Consolidated Transactions"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optAutoGenSig 
            BackColor       =   &H80000018&
            Caption         =   "Single Transactions"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   1815
         End
         Begin HoverButton.HoverControl cmdAutoDmdGenCalcel 
            Height          =   255
            Left            =   1920
            TabIndex        =   77
            Top             =   1320
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   -2147483647
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Close"
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "->>Automatic Demand Generate:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   195
            Left            =   20
            TabIndex        =   32
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   580
         Left            =   9480
         TabIndex        =   63
         Top             =   4560
         Width           =   3615
         Begin VB.CommandButton cmdPrintThis 
            BackColor       =   &H00FAD5DF&
            Caption         =   "Print Selected"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FAD5DF&
            Caption         =   "Print All"
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintAll 
            BackColor       =   &H00FAD5DF&
            Caption         =   "&All Unprinted"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame fraPrintChoice 
         BackColor       =   &H00D7EBD8&
         Height          =   7575
         Left            =   2400
         TabIndex        =   43
         Top             =   7560
         Visible         =   0   'False
         Width           =   11415
         Begin VB.Frame fraChoice1 
            BackColor       =   &H00EFEFDF&
            Caption         =   "Choice 1:"
            Height          =   1095
            Left            =   240
            TabIndex        =   58
            Top             =   240
            Width           =   2295
            Begin VB.OptionButton optInvoice 
               BackColor       =   &H00EFEFDF&
               Caption         =   "Invloice"
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optReport 
               BackColor       =   &H00EFEFDF&
               Caption         =   "Report"
               Height          =   375
               Left            =   120
               TabIndex        =   59
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.Frame fraChoice3 
            BackColor       =   &H00EFEFDF&
            Caption         =   "Choice 3:"
            Enabled         =   0   'False
            Height          =   1095
            Left            =   5880
            TabIndex        =   55
            Top             =   240
            Width           =   2775
            Begin VB.OptionButton optSingle 
               BackColor       =   &H00EFEFDF&
               Caption         =   "Single Transaction"
               Height          =   375
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optAll 
               BackColor       =   &H00EFEFDF&
               Caption         =   "All Transactions"
               Height          =   375
               Left            =   120
               TabIndex        =   56
               Top             =   600
               Width           =   1695
            End
         End
         Begin VB.Frame fraChoice2 
            BackColor       =   &H00EFEFDF&
            Caption         =   "Choice 2:"
            Enabled         =   0   'False
            Height          =   1095
            Left            =   2880
            TabIndex        =   51
            Top             =   240
            Width           =   2415
            Begin VB.OptionButton optPayeeWise 
               BackColor       =   &H00EFEFDF&
               Caption         =   "Payee wise"
               Height          =   375
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton optBatchWise 
               BackColor       =   &H00EFEFDF&
               Caption         =   "Batch wise"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   720
               Width           =   1215
            End
            Begin VB.ComboBox cmbBatchWise 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1320
               TabIndex        =   52
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "New Demand"
            Height          =   1215
            Left            =   8400
            TabIndex        =   48
            Top             =   6240
            Width           =   2895
            Begin MSForms.CommandButton cmdDemandPrint 
               Height          =   975
               Left            =   240
               TabIndex        =   50
               Top             =   120
               Width           =   975
               Caption         =   "Print"
               PicturePosition =   131080
               Size            =   "1720;1720"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdDemandCancel 
               Height          =   975
               Left            =   1680
               TabIndex        =   49
               Top             =   120
               Width           =   975
               Caption         =   "Cancel"
               PicturePosition =   65543
               Size            =   "1720;1720"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1215
            Left            =   120
            TabIndex        =   44
            Top             =   6240
            Width           =   3135
            Begin VB.CheckBox chkMarkPrinted 
               BackColor       =   &H00FBC9CD&
               Caption         =   "Mark as printed"
               Height          =   375
               Left            =   120
               TabIndex        =   46
               Top             =   720
               Value           =   1  'Checked
               Width           =   2535
            End
            Begin VB.TextBox txtPrintedRef 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Print Ref:"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   360
               Width           =   660
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   4215
            Left            =   120
            TabIndex        =   61
            Top             =   1800
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   7435
            _Version        =   393216
            Cols            =   11
            BackColorBkg    =   15325144
            _NumberOfBands  =   1
            _Band(0).Cols   =   11
         End
         Begin MSForms.CommandButton CommandButton1 
            Default         =   -1  'True
            Height          =   1095
            Left            =   9360
            TabIndex        =   62
            Top             =   240
            Width           =   1215
            VariousPropertyBits=   19
            Caption         =   "GO"
            PicturePosition =   393224
            Size            =   "2143;1931"
            FontName        =   "MS Serif"
            FontEffects     =   1073741825
            FontHeight      =   270
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame fraDetails 
         BackColor       =   &H00F3F6FD&
         Caption         =   "Demand Details: "
         Height          =   2775
         Left            =   120
         TabIndex        =   41
         Top             =   5160
         Width           =   13455
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChildDemands 
            Height          =   2415
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   4260
            _Version        =   393216
            BackColor       =   15988477
            BackColorFixed  =   13883619
            BackColorBkg    =   13883619
            GridColor       =   -2147483635
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraMain 
         Height          =   4800
         Left            =   80
         TabIndex        =   10
         Top             =   420
         Width           =   1140
         Begin VB.Frame fraDeleteDemand 
            BackColor       =   &H80000010&
            Height          =   4695
            Left            =   0
            TabIndex        =   11
            Top             =   4080
            Visible         =   0   'False
            Width           =   1395
            Begin VB.CommandButton cmdDeleteOld 
               BackColor       =   &H80000003&
               Caption         =   "Old Demands"
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   960
               Width           =   1200
            End
            Begin VB.CommandButton cmdDelete 
               BackColor       =   &H80000003&
               Caption         =   "Demands"
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label lblDeleteDemands 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               Caption         =   "Delete"
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
               Height          =   375
               Left            =   0
               MousePointer    =   99  'Custom
               TabIndex        =   14
               Top             =   0
               Width           =   1400
            End
         End
         Begin VB.Frame fraReprintDemands 
            BackColor       =   &H80000010&
            Height          =   2895
            Left            =   0
            TabIndex        =   18
            Top             =   3480
            Visible         =   0   'False
            Width           =   1395
            Begin VB.CommandButton cmdReprintAll 
               BackColor       =   &H00E5D5C5&
               Caption         =   "&All"
               Height          =   375
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   960
               Width           =   1200
            End
            Begin VB.CommandButton cmdReprintSome 
               BackColor       =   &H00E5D5C5&
               Caption         =   "&Selected"
               Height          =   375
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   1440
               Width           =   1200
            End
            Begin VB.CommandButton cmdCancelReprint 
               BackColor       =   &H00E5D5C5&
               Caption         =   "&Cancel"
               Height          =   375
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   1920
               Width           =   1200
            End
            Begin VB.CommandButton cmdReprint 
               BackColor       =   &H00E5D5C5&
               Caption         =   "&Reprint"
               Height          =   375
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label lblReprintDemand 
               Alignment       =   2  'Center
               BackColor       =   &H00E5D5C5&
               Caption         =   "Reprint"
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
               Height          =   375
               Left            =   0
               MousePointer    =   99  'Custom
               TabIndex        =   23
               Top             =   0
               Width           =   1395
            End
         End
         Begin VB.Frame fraEditDemand 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Height          =   2415
            Left            =   0
            TabIndex        =   15
            Top             =   480
            Width           =   1155
            Begin VB.CommandButton cmdEdit 
               BackColor       =   &H00E5E5E5&
               Caption         =   "&Edit"
               Height          =   495
               Left            =   50
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   600
               Width           =   1080
            End
            Begin VB.Label lblEditDemand 
               Alignment       =   2  'Center
               BackColor       =   &H00E5E5E5&
               Caption         =   "Edit Demands"
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
               Height          =   375
               Left            =   0
               MousePointer    =   99  'Custom
               TabIndex        =   17
               Top             =   0
               Width           =   1155
            End
         End
         Begin VB.Frame fraGenerate 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Height          =   6615
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   1155
            Begin VB.CommandButton cmdPostDemands 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Post Demands"
               Height          =   495
               Left            =   50
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   1440
               Width           =   1080
            End
            Begin VB.CommandButton cmdGenAll 
               BackColor       =   &H00C0E0FF&
               Caption         =   "&Automatic..."
               Height          =   495
               Left            =   50
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   480
               Width           =   1080
            End
            Begin VB.CommandButton cmdClearDemands 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Clear"
               Height          =   495
               Left            =   50
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   1920
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CommandButton cmdGenerateManual 
               BackColor       =   &H00C0E0FF&
               Caption         =   "&Manual"
               Height          =   495
               Left            =   50
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label lblGenerate 
               Alignment       =   2  'Center
               BackColor       =   &H00C0E0FF&
               Caption         =   "Generate Demands"
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
               Height          =   375
               Left            =   0
               MousePointer    =   99  'Custom
               TabIndex        =   28
               Top             =   0
               Width           =   1155
            End
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemands 
         Height          =   4095
         Left            =   1320
         TabIndex        =   33
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   7223
         _Version        =   393216
         GridColor       =   -2147483635
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CheckBox chkSelectAllDemands 
         Caption         =   "&Select All"
         Height          =   255
         Left            =   1560
         TabIndex        =   287
         Top             =   4680
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   600
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmDemands1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bAddNew As Boolean
Dim iSelectedDemandsRow As Integer, iIncDec As Integer
Dim szaFreq() As String, szPrefix As String
Dim szLastIDClicked As String
Dim szCurCompName As String, szCurCompSageAccNum As String
Dim objDemandType As clsArray
Dim fVAT_Rate As Single
Dim szAllBankBalance As String
Dim iCurRow As Integer
Dim bChangesMade As Boolean
Dim baChangesMade() As Boolean
Private sTextBox As String
Public nTaxCode As Double
Private iSelected As Integer
Private iCurEditRow As Integer
Private curOpeningBal As Currency

Public TenantCode As String
Public TenantName As String
Public Unit As String
Public typeofdemand As Integer
Public text As String
Dim szIC As String
Public Amount As Double

Private Sub cboEditType_Click()
   Dim szTemp() As String, DemandType As String

   szTemp = Split(cboEditType.text, " / ")

   Nominal CInt(szTemp(0)), flxEditDemand.Row, flxEditDemand

   flxEditDemand.TextMatrix(flxEditDemand.Row, 17) = szPrefix & Format(txtEditIssueDate.text, "ddmmyy")

   flxEditDemand.TextMatrix(flxEditDemand.Row, 2) = szTemp(1)
   cboEditType.Visible = False
End Sub

Private Sub cboEditType_LostFocus()
'   if cboedittype.Text<>"" then
End Sub

Private Sub cboEditVatCode_Click()
   Dim szTemp() As String

   szTemp = Split(cboEditVatCode.text, " / ")

   flxEditDemand.TextMatrix(flxEditDemand.Row, 18) = szTemp(0)
   fVAT_Rate = CSng(szTemp(2))
   If flxEditDemand.TextMatrix(flxEditDemand.Row, 8) <> "" Then
      flxEditDemand.TextMatrix(flxEditDemand.Row, 9) = _
                        Format(CCur(flxEditDemand.TextMatrix(flxEditDemand.Row, 8)) * _
                        (fVAT_Rate / 100), "0.00")
      flxEditDemand.TextMatrix(flxEditDemand.Row, 10) = _
                        Format(CCur(flxEditDemand.TextMatrix(flxEditDemand.Row, 8)) + _
                        CCur(flxEditDemand.TextMatrix(flxEditDemand.Row, 9)), "0.00")
      CalSubTotal flxEditDemand, txtEditSubAmount, txtEditSubVat, txtEditSubTotal
   End If

   cboEditVatCode.Visible = False
End Sub

Private Sub cboEditVatCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then cboEditVatCode.Visible = False
End Sub

Private Sub cboEditVatCode_LostFocus()
   cmdUpdateDemand.SetFocus
End Sub

Private Sub cboTenant_Click()
   Dim szaComp() As String
'
   szaComp = Split(cboTenant.text, " / ")
   txtAddNewSageText = "S/L " & szaComp(0)
   szCurCompSageAccNum = szaComp(0)
   szCurCompName = szaComp(1)
   txtUnit.text = GetUnitIDbyTenantID(szCurCompSageAccNum)
'
   txtAddNewSageText.text = "S/L " & szaComp(0)
'
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 3) = "M"
'
   txtTenantName.text = cboTenant.text
   cboTenant.Visible = False
End Sub

Private Sub cboTenant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then cboTenant.Visible = False
End Sub

Private Sub cboType_Click()
   Dim szTemp() As String, DemandType As String
'
   szTemp = Split(cboType.text, " / ")
   objDemandType.AddItemPos szTemp(1), CInt(szTemp(0))
'
   Nominal CInt(szTemp(0)), flxAddNewDemands.Rows - 1, flxAddNewDemands
'
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 17) = szPrefix & Format(txtIssueDate.text, "ddmmyy")
'
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 2) = szTemp(1)
'   cboType.Visible = False
End Sub

Private Sub Nominal(DTId As Integer, iFlxRow As Integer, conFlxGrid As Control)
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String
'
'   connect to database
   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt
'
'   Get the details for the demand type selected
   SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = " & DTId & ""
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
   conFlxGrid.TextMatrix(iFlxRow, 11) = IIf(IsNull(rdoRst1!NominalCodeforAmount), "", rdoRst1!NominalCodeforAmount)
   conFlxGrid.TextMatrix(iFlxRow, 12) = IIf(IsNull(rdoRst1!NominalNameforAmount), "", rdoRst1!NominalNameforAmount)
   conFlxGrid.TextMatrix(iFlxRow, 13) = IIf(IsNull(rdoRst1!NominalCodeForVAT), "", rdoRst1!NominalCodeForVAT)
   conFlxGrid.TextMatrix(iFlxRow, 14) = IIf(IsNull(rdoRst1!NominalNameforVAT), "", rdoRst1!NominalNameforVAT)
   conFlxGrid.TextMatrix(iFlxRow, 15) = IIf(IsNull(rdoRst1!NominalCodeForTotal), "", rdoRst1!NominalCodeForTotal)
   conFlxGrid.TextMatrix(iFlxRow, 16) = IIf(IsNull(rdoRst1!NominalNameforTotal), "", rdoRst1!NominalNameforTotal)
'
   szPrefix = rdoRst1!Prefix
   szIC = rdoRst1!InvCrd
   rdoRst1.Close
   rdoConn.Close
'
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Sub

Private Sub cmbManualBatchNo_Click()
'   If cmbManualBatchNo.text = "NEW" Then
'      cmbManualBatchNo.AddItem CreateTransactionID("B", 1), 1
'      cmbManualBatchNo.ListIndex = 1
'   End If
'   If cmbManualBatchNo.text = "NONE" Then
'      cmbManualBatchNo.text = ""
'   End If
End Sub

Private Sub cboType_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cboType.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then cboType.Visible = False
End Sub

Private Sub cboType_LostFocus()
   cboType.Visible = False
   flxAddNewDemands.ColSel = 4         'Description
   flxAddNewDemands_Click
End Sub

Private Sub cboVatCode_Click()
   Dim szTemp() As String
'
   szTemp = Split(cboVatCode.text, " / ")
'
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 18) = szTemp(0)
   fVAT_Rate = CSng(szTemp(2))
   If flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 8) <> "" Then
      flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 9) = _
                        Format(CCur(flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 8)) * _
                        (fVAT_Rate / 100), "0.00")
      flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 10) = _
                        Format(CCur(flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 8)) + _
                        CCur(flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 9)), "0.00")
      CalSubTotal flxAddNewDemands, txtSubTAmount, txtSubTVAT, txtSubTTotal
   End If
'
   cboVatCode.Visible = False
End Sub

Private Sub cboVatCode_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cboVatCode.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboVatCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then cboVatCode.Visible = False
End Sub

Private Sub chkOldTenats_Change()
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String

   cboTenant.Clear

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   If chkOldTenats.Value Then
      SQLStr1 = "SELECT CompanyName, Tenants.SageAccountNumber " & _
                "FROM Tenants " & _
                "ORDER BY Tenants.SageAccountNumber"
   Else
      SQLStr1 = "SELECT CompanyName, Tenants.SageAccountNumber " & _
                "FROM Tenants, Units " & _
                "WHERE Tenants.SageAccountNumber = Units.SageAccountNumber AND " & _
                      "Units.Occupied = 'Y' " & _
                "ORDER BY Tenants.SageAccountNumber"
   End If
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   While Not rdoRst1.EOF
       cboTenant.AddItem rdoRst1!SageAccountNumber & " / " & rdoRst1!CompanyName
       rdoRst1.MoveNext
   Wend

   rdoRst1.Close
   rdoConn.Close
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Sub

Private Sub chkSelectAllDemands_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 32 Then Exit Sub
   Dim iRow As Integer, i As Integer
'
   If chkSelectAllDemands.Value Then
      For i = 1 To flxDemands.Rows - 1
         flxDemands.TextMatrix(i, 0) = ""
      Next i
      For i = 1 To flxDemands.Rows - 1
         iIncDec = SelectFlxGridRow(flxDemands, i)
      Next i
      FlxGridConfigure flxChildDemands
   '
      szLastIDClicked = flxDemands.TextMatrix(flxDemands.Rows - 1, 1)
      szCurCompName = flxDemands.TextMatrix(flxDemands.Rows - 1, 2)
      Call FillChildinGrid(szLastIDClicked, flxChildDemands)
      fraDetails.Caption = "Demand Details: " & szLastIDClicked
      iSelectedDemandsRow = flxDemands.Rows - 1
   Else
      For i = 1 To flxDemands.Rows - 1
         iIncDec = SelectFlxGridRow(flxDemands, i)
      Next i
      FlxGridConfigure flxChildDemands
      szLastIDClicked = ""
      fraDetails.Caption = "Demand Details: " & szLastIDClicked
      iSelectedDemandsRow = 0
   End If
End Sub

Private Sub chkSelectAllDemands_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim iRow As Integer, i As Integer
'
   If chkSelectAllDemands.Value Then
      For i = 1 To flxDemands.Rows - 1
         flxDemands.TextMatrix(i, 0) = ""
      Next i
      For i = 1 To flxDemands.Rows - 1
         iIncDec = SelectFlxGridRow(flxDemands, i)
      Next i
      FlxGridConfigure flxChildDemands
   '
      szLastIDClicked = flxDemands.TextMatrix(flxDemands.Rows - 1, 1)
      szCurCompName = flxDemands.TextMatrix(flxDemands.Rows - 1, 2)
      Call FillChildinGrid(szLastIDClicked, flxChildDemands)
      fraDetails.Caption = "Demand Details: " & szLastIDClicked
      iSelectedDemandsRow = flxDemands.Rows - 1
   Else
      For i = 1 To flxDemands.Rows - 1
         iIncDec = SelectFlxGridRow(flxDemands, i)
      Next i
      FlxGridConfigure flxChildDemands
      szLastIDClicked = ""
      fraDetails.Caption = "Demand Details: " & szLastIDClicked
      iSelectedDemandsRow = 0
   End If
End Sub

Private Sub cmbBankAc_Click()
   Dim szaBankBal() As String
   
   szaBankBal = Split(szAllBankBalance, " # ")
   
   txtChqNo.text = Format(szaBankBal(cmbBankAc.ListIndex), "0.00")
   curOpeningBal = CCur(txtChqNo.text)
End Sub

Private Sub cmbBankAc_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cmbBankAc.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbTenant_Click()
   SPFlxGridConfigure
   LoadDataInGrid
   
   ReDim baChangesMade(flxSPayment.Rows) As Boolean

End Sub

Private Sub cmbTenant_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cmbTenant.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmdAddNewDemand_Click()
   If txtTenantName.text = "" Then
      MsgBox "Please select the tenant.", vbCritical + vbOKOnly, "Tenant Selection"
      Exit Sub
   End If
'
   If txtIssueDate.text = "" Then
      MsgBox "Plseas select due date.", vbCritical + vbOKOnly, "Due Date"
      txtIssueDate_Click
      Exit Sub
   End If
'
   If flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 4) = "" Then
      MsgBox "Please give the description of the last statement.", vbCritical + vbOKOnly, "Error"
      flxAddNewDemands.Col = 4
      flxAddNewDemands_Click
      Exit Sub
   End If
'
   If flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 18) = "" And flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 8) <> "" Then
      MsgBox "Please give the VAT Code.", vbCritical + vbOKOnly, "Error"
      flxAddNewDemands.Col = 18
      flxAddNewDemands_Click
      Exit Sub
   End If
'
   Dim iTempFlxRow As Integer, iRow As Integer, iCol As Integer
'I need to make sure its needed
'   If bAddNew Then
'      flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 8) = LTrim(Right(flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 8), Len(flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 8)) - 1))
'      CalSubTotal flxAddNewDemands, txtSubTAmount, txtSubTVAT, txtSubTTotal
'   End If
'
   flxAddNewDemands.AddItem ""         'Add a blank line at the bottom of the flxGrid
   txtAddNewDescription.text = ""
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 1) = Format(CInt(flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 2, 1)) + 1, "000")
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 3) = "M"
'
   bAddNew = True
End Sub

Private Sub cmdAutoDmdGenCalcel_Click()
   fraAutoDemandChoice.Visible = False
   fraMain.Enabled = True
End Sub

Private Sub cmdAutoDmdGenOk_Click()
   fraAutoDemandChoice.Visible = False
   fraMain.Enabled = True

   If optAutoGenConsolidated.Value = True Then
      Call GenAutoConDemands
   Else
      Call GenAutoSngDemands
   End If
   FlxDemandsConfigure flxDemands
   FillDemandsFlxGrid flxDemands, False
   flxDemands.Row = 0
   flxDemands.Col = 0
End Sub

Private Sub cmdBkList_Click(Index As Integer)
   LoadBankAccount
'
   fraListBk(tabPayment.Tab - 1).Left = txtBkAc(tabPayment.Tab - 1).Left + fraBkInput(tabPayment.Tab - 1).Left
   fraListBk(tabPayment.Tab - 1).Top = txtBkAc(tabPayment.Tab - 1).Top + txtBkAc(tabPayment.Tab - 1).Height + _
                                    fraBkInput(tabPayment.Tab - 1).Top
   fraListBk(tabPayment.Tab - 1).Visible = True
   fraListBk(tabPayment.Tab - 1).ZOrder 0
   flxListBk(tabPayment.Tab - 1).SetFocus
   sTextBox = "Bank"
End Sub

Private Sub LoadBankAccount()
   flxListBk(tabPayment.Tab - 1).TextMatrix(0, 0) = "Bank Code"
   flxListBk(tabPayment.Tab - 1).TextMatrix(0, 1) = "Bank Name"
   flxListBk(tabPayment.Tab - 1).ColWidth(0) = 800
   flxListBk(tabPayment.Tab - 1).ColWidth(1) = 2500
   ' Error Handler
   On Error GoTo Error_Handler
   
   Dim clsBankAC As clsArray
   Dim iBankAc As Integer
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oBankRecord As SageDataObject120.BankRecord
   Dim oNominalRecord As SageDataObject120.NominalRecord

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
   
      Set oBankRecord = oWS.CreateObject("BankRecord")
   
      ' Move to the First Record
      oBankRecord.MoveFirst
      Set clsBankAC = New clsArray
      For iBankAc = 1 To oBankRecord.Count
         clsBankAC.AddItem oBankRecord.Fields.Item("ACCOUNT_REF").Value
         oBankRecord.MoveNext
      Next iBankAc
   
      Set oBankRecord = Nothing
      
      Set oNominalRecord = oWS.CreateObject("NominalRecord")
   
      oNominalRecord.MoveFirst
      
      Dim rRow As Integer
      Dim iRec As Integer
      rRow = 1
      For iRec = 1 To oNominalRecord.Count
         If clsBankAC.IsItem(CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)) Then
            flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 0) = CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)
            flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 1) = CStr(oNominalRecord.Fields.Item("NAME").Value)
            flxListBk(tabPayment.Tab - 1).AddItem ""
            rRow = rRow + 1
         End If
         oNominalRecord.MoveNext
      Next iRec
      'Disconnect
      oWS.Disconnect
   End If

   ' Destroy Objects
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text
   
   Set oBankRecord = Nothing
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub cmdCancelBk_Click(Index As Integer)
   Dim iRow As Integer, iRemRec As Integer
'
   If Not cmdEditBk(tabPayment.Tab - 1).Enabled Then
      If MsgBox("Do you want to cancel Edit?", vbQuestion + vbYesNo, "Edit Record") = vbNo Then Exit Sub
      HandleTextBoxesBk True, False
      iSelected = iSelected + Select1RowFlxGrid(flxBankPay(tabPayment.Tab - 1), _
                  iCurEditRow, flxBankPay(tabPayment.Tab - 1).Cols - 2)
      cmdEditBk(tabPayment.Tab - 1).Enabled = True
   End If
   If Not cmdNewBk(tabPayment.Tab - 1).Enabled Then
      If MsgBox("Do you want to cancel the new records?", vbQuestion + vbYesNo, "Add Record") = vbNo Then Exit Sub
      
      iRemRec = 0
      iRow = 1
      While iRow <= flxBankPay(tabPayment.Tab - 1).Rows - 1
         If flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 14) = "" Then
            iRemRec = iRemRec + 1
            flxBankPay(tabPayment.Tab - 1).RemoveItem (iRow)
            iRow = iRow - 1
         End If
         iRow = iRow + 1
      Wend
'
      ReduceLastID "tlbBankPayment", iRemRec
'
      HandleTextBoxesBk True, False
      cmdNewBk(tabPayment.Tab - 1).Enabled = True
      bChangesMade = False
   End If

End Sub

Private Sub cmdCancelNew_Click()
   If MsgBox("Are you sure to discard all new data?", vbYesNo + vbQuestion, "Add New Demands") = vbNo Then Exit Sub
   fraCreateManualDemand.Visible = False
   MsgBox "The manual demand has not been saved.", vbOKOnly + vbInformation, "Cancelled"
   fraMain.Enabled = True
   fraDetails.Visible = True
End Sub

Private Sub cmdCancelPrint_Click()
''The user has selected to exit Print Mode
'Set adoConn = New ADODB.Connection
'Set adoRst = New ADODB.Recordset
'
''connect to database
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
''get sendtoprint field for all demands that were set to be sent to print
'strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
'adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
''set the sendtoprint field to not be sent to print - blank
'If adoRst.EOF = False Then
'    While adoRst.EOF = False
'        adoRst!SendToPrint = ""
'        adoRst.Update
'        adoRst.MoveNext
'    Wend
'End If
'adoRst.Close
'adoConn.Close
'Set adoRst = Nothing
'Set adoConn = Nothing
'
'cmdCancelPrint.Visible = False
'cmdPrintAll.Visible = False
'cmdPrintBatch.Visible = True
'cmdPrintSome.Visible = False
'cmdGenAll.Visible = True
'cmdGenerateManual.Visible = True
'cmdEdit.Visible = True
'cmdDelete.Visible = True
'cmdDeleteOld.Visible = True
'cmdPrint.Visible = True
'cmdPrintThis.Visible = True
'cmdReprint.Visible = True
'chkPrint.Visible = False
''lbl1.Visible = False
'Call EnableMenu
'
'PrintMode = False
'
'Call EmptyBoxes
'Call GetFirstDemand

End Sub

Private Sub cmdCancelReprint_Click()
''user selected to exit Reprint mode
'Set adoConn = New ADODB.Connection
'Set adoRst = New ADODB.Recordset
'
''connect to database
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
''get the sendtoprint field from demands that were set to be sent to print
'strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
'adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
''set the sendtoprint field to not be sent to print - blank
'If adoRst.EOF = False Then
'    While adoRst.EOF = False
'        adoRst!SendToPrint = ""
'        adoRst.Update
'        adoRst.MoveNext
'    Wend
'End If
'adoRst.Close
'adoConn.Close
'Set adoRst = Nothing
'Set adoConn = Nothing
'
'ReprintMode = False
'
'cmdCancelReprint.Visible = False
'cmdReprintSome.Visible = False
'cmdReprintAll.Visible = False
'cmdPrintBatch.Visible = False
'cmdPrintAll.Visible = False
'cmdPrintBatch.Visible = True
'cmdPrintSome.Visible = False
'cmdGenAll.Visible = True
'cmdGenerateManual.Visible = True
'cmdEdit.Visible = True
'cmdDelete.Visible = True
'cmdDeleteOld.Visible = True
'cmdPrint.Visible = True
'cmdPrintThis.Visible = True
'cmdReprint.Visible = True
'chkPrint.Visible = False
''lbl1.Visible = False
'Call EnableMenu
'
'Call EmptyBoxes
'Call GetFirstDemand

End Sub

Private Sub cmdCancelUpdate_Click()
   If MsgBox("Are you sure to discard all changes?", vbYesNo + vbQuestion, "Edit Demands") = vbNo Then Exit Sub
   fraEditDemandWindow.Visible = False
   MsgBox "The demand has not been changed!", vbOKOnly + vbInformation, "Demand"
'
   cmdEdit.Enabled = True
   fraDetails.Visible = True
End Sub

Private Sub cmdCCBk_Click(Index As Integer)
   LoadCCBk
'
   fraListBk(tabPayment.Tab - 1).Left = txtCCBk(tabPayment.Tab - 1).Left + fraBkInput(tabPayment.Tab - 1).Left
   fraListBk(tabPayment.Tab - 1).Top = txtCCBk(tabPayment.Tab - 1).Top + txtCCBk(tabPayment.Tab - 1).Height + fraBkInput(tabPayment.Tab - 1).Top
   fraListBk(tabPayment.Tab - 1).Visible = True
   fraListBk(tabPayment.Tab - 1).ZOrder 0
   flxListBk(tabPayment.Tab - 1).SetFocus
   sTextBox = "CC"
End Sub

Private Function LoadCCBk() As Boolean
   flxListBk(tabPayment.Tab - 1).ColWidth(0) = 1500
   flxListBk(tabPayment.Tab - 1).ColWidth(1) = 2700
   flxListBk(tabPayment.Tab - 1).ColAlignment = vbLeftJustify
   
   ' Error Handler
   On Error GoTo Error_Handler
   
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oProjCostCodes As SageDataObject120.ProjectCostCodes

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
   
      ' Create Objects
      Set oProjCostCodes = oWS.CreateProjectCostCodes
      
      LoadCCBk = True
      If oProjCostCodes.Count = 0 Then
         MsgBox "No Cost Code has been created in SAGE", vbCritical, "Cost Code Empty"
         LoadCCBk = False
         GoTo Error_Handler
      End If
      
      flxListBk(tabPayment.Tab - 1).Clear
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 0) = "CODE"
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 1) = "DESCRIPTION"
      
      Dim rRow As Integer
      For rRow = 1 To oProjCostCodes.Count
         flxListBk(tabPayment.Tab - 1).AddItem ""
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 0) = CStr(oProjCostCodes.Item(rRow - 1).Reference)
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 1) = CStr(oProjCostCodes.Item(rRow - 1).description)
      Next rRow
      
      flxListBk(tabPayment.Tab - 1).RemoveItem oProjCostCodes.Count + 1
      flxListBk(tabPayment.Tab - 1).RemoveItem oProjCostCodes.Count + 1
   
      'Disconnect
      oWS.Disconnect
   End If

   ' Destroy Objects
   Set oProjCostCodes = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Function

   ' Error Handling Code
Error_Handler:

   MsgBox "(pcm_008) The SDO generated the following error: " & oSDO.LastError.text
   
   Set oProjCostCodes = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Function

Private Sub cmdCloseBk_Click(Index As Integer)
   If Not cmdNewBk(tabPayment.Tab - 1).Enabled Or Not cmdEditBk(tabPayment.Tab - 1).Enabled Or bChangesMade Then
      If MsgBox("You want to close this window? Your data may be lost.", vbInformation + vbYesNo, "Close this window") = vbNo Then Exit Sub
   End If
   Unload Me
End Sub

Private Sub cmdCloseBk_LostFocus(Index As Integer)
   tabPayment.SetFocus
End Sub

Private Sub cmdDeleteOld_Click()
'user wants to delete all old demands - ones that have been printed, exported to sage and exported to excel
If MsgBox("Do you really want to delete old demands?", vbYesNo + vbQuestion, "Delete Old Demands") = vbNo Then Exit Sub
MousePointer = vbHourglass

Call DeleteDemands

Call EmptyBoxes
Call GetFirstDemand
MsgBox "Old demands deleted successfully", vbOKOnly + vbInformation, "Deleted"
MousePointer = vbDefault

End Sub

Private Sub cmdDemandCancel_Click()
   If MsgBox("Do you want to cancel?", vbQuestion + vbYesNo, "Print Cancel") = vbYes Then
      fraPrintChoice.Visible = False
   End If
End Sub

Private Sub cmdDemandPrint_Click()
   If MsgBox("Do you want to mark these transactions as printed?", vbYesNo + vbQuestion, "Mark as printed") = vbYes Then
      chkMarkPrinted.Value = vbChecked
   Else
      chkMarkPrinted.Value = vbUnchecked
   End If
End Sub

Private Sub cmdDeptBk_Click(Index As Integer)
   MousePointer = vbHourglass
   LoadDeptBk
'
   fraListBk(tabPayment.Tab - 1).Left = txtDeptBk(tabPayment.Tab - 1).Left + fraBkInput(tabPayment.Tab - 1).Left
   fraListBk(tabPayment.Tab - 1).Top = txtDeptBk(tabPayment.Tab - 1).Top + txtDeptBk(tabPayment.Tab - 1).Height + fraBkInput(tabPayment.Tab - 1).Top
   fraListBk(tabPayment.Tab - 1).Visible = True
   fraListBk(tabPayment.Tab - 1).ZOrder 0
   flxListBk(tabPayment.Tab - 1).SetFocus
   sTextBox = "Dept"
   MousePointer = vbDefault
End Sub

Private Sub LoadDeptBk()
   flxListBk(tabPayment.Tab - 1).ColWidth(0) = 1500
   flxListBk(tabPayment.Tab - 1).ColWidth(1) = 2700
   flxListBk(tabPayment.Tab - 1).ColAlignment = vbLeftJustify

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
      flxListBk(tabPayment.Tab - 1).Clear
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 0) = "Dept. ID"
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 1) = "Department Name"
      
      Dim rRow As Integer
      For rRow = 1 To oDepartmentData.Count
         oDepartmentData.Read (rRow)
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 0) = CStr(rRow)
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 1) = CStr(oDepartmentData.Fields.Item("NAME").Value)
         flxListBk(tabPayment.Tab - 1).AddItem ""
      Next rRow
      'Disconnect
      oWS.Disconnect
   End If

   ' Destroy Objects
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "(frmDemands1 LoadDeptBk) The SDO generated the following error: " & oSDO.LastError.text
   
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub cmdEdit_Click()
   If iIncDec <> 1 Then
      MsgBox "Please select one demand only.", vbInformation + vbOKOnly, "Demand Selection"
      Exit Sub
   End If
   If flxDemands.TextMatrix(flxDemands.Row, 9) = "YES" Then
      MsgBox "After exported to SAGE Demand you cannot edit!", vbInformation + vbOKOnly, "Edit Demand"
      Exit Sub
   End If
'
   fraEditDemandWindow.Left = flxDemands.Left
   fraEditDemandWindow.Top = flxDemands.Top
   fraEditDemandWindow.Visible = True
   fraEditDemandWindow.ZOrder 0
'
   Call Edit
'
   cmdEdit.Enabled = False
   fraDetails.Visible = False
End Sub

Private Sub cmdFind_Click()
   'user wants to find a demand
   Dim i As Integer
   Dim char As String, strSQLTitles As String
   Dim check As Boolean
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
'
'   If Text2.text = "" Then
'       MsgBox "You must select the Demand ID of the demand you want to find!", vbOKOnly + vbCritical, "Find Demand"
'       Exit Sub
'   End If
   'check that user entered a number in the text box
   check = True
'   For i = 1 To Len(Text2.text)
'       char = Mid(Text2.text, i, 1)
'       If Asc(char) < 48 Or Asc(char) > 57 Then check = False
'   Next i
'
   If check = False Then
       MsgBox "Invalid Demand Reference Number!", vbOKOnly + vbCritical, "Invalid Number"
       Exit Sub
   End If
'
   Set adoConn = New ADODB.Connection
   Set adoRst = New ADODB.Recordset
   'connect to database
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   'get the details of the demand with uniquerefnumber the user entered
'   strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & Text2.text & "';"
   'for the various modes only select it if the printed and exported to sage status is ok.
'   If EditMode = True Then strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & Text2.text & "' AND UPDATE_SAGE = 'N'"
'   If PrintMode = True Then strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & Text2.text & "' AND IsPrinted = 'N'"
'   If ReprintMode = True Then strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & Text2.text & "' AND IsPrinted = 'Y'"
'
   adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
   'if record exists display on screen if not tell user invalid Demand Id
   If adoRst.EOF = False Then
       Call EmptyBoxes
       Call GetRecord
   Else
       MsgBox "You have entered an invalid Demand ID", vbOKOnly + vbCritical, "Invalid Demand Id"
   End If
   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
'
'   Text2.text = ""
End Sub

Private Sub cmdEditBk_Click(Index As Integer)
   If cmdNewBk(tabPayment.Tab - 1).Enabled = False Then Exit Sub
   If iSelected = 0 Then
      MsgBox "Select at least 1 row.", vbInformation + vbOKOnly, "Edit Record"
      Exit Sub
   End If
'
   HandleTextBoxesBk True, True
'
   txtBkAc(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 0)
   txtDateBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 1)
   txtTypeBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 2)
   txtUnitBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 4)
   txtNCBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 5)
   txtDeptBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 6)
   txtProjBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 7)
   txtCCBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 8)
   txtDetailsBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 9)
   txtNetBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 10)
   cmdTaxListBk(tabPayment.Tab - 1).Caption = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 11)
   txtVatBk(tabPayment.Tab - 1).text = flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 12)
   txtRecharge(tabPayment.Tab - 1).text = IIf(flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 13) = "X", "YES", "NO")
'
   flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, 16) = "1"
'
   iCurEditRow = flxBankPay(tabPayment.Tab - 1).Row
   cmdEditBk(tabPayment.Tab - 1).Enabled = False
   bChangesMade = True
End Sub

Private Sub HandleTextBoxesBk(bClear As Boolean, bEnable As Boolean)
   If bClear Then
      txtBkAc(tabPayment.Tab - 1).text = ""
      txtDateBk(tabPayment.Tab - 1).text = "//"
      txtTypeBk(tabPayment.Tab - 1).text = ""
      txtUnitBk(tabPayment.Tab - 1).text = ""
      txtNCBk(tabPayment.Tab - 1).text = ""
      txtDeptBk(tabPayment.Tab - 1).text = ""
      txtProjBk(tabPayment.Tab - 1).text = ""
      txtCCBk(tabPayment.Tab - 1).text = ""
      txtDetailsBk(tabPayment.Tab - 1).text = ""
      txtNetBk(tabPayment.Tab - 1).text = ""
      txtVatBk(tabPayment.Tab - 1).text = ""
      txtRecharge(tabPayment.Tab - 1).text = "NO"
   End If
      txtBkAc(tabPayment.Tab - 1).Enabled = bEnable
      txtDateBk(tabPayment.Tab - 1).Enabled = bEnable
      txtTypeBk(tabPayment.Tab - 1).Enabled = bEnable
      txtUnitBk(tabPayment.Tab - 1).Enabled = bEnable
      txtNCBk(tabPayment.Tab - 1).Enabled = bEnable
      txtDeptBk(tabPayment.Tab - 1).Enabled = bEnable
      txtProjBk(tabPayment.Tab - 1).Enabled = bEnable
      txtCCBk(tabPayment.Tab - 1).Enabled = bEnable
      txtDetailsBk(tabPayment.Tab - 1).Enabled = bEnable
      txtNetBk(tabPayment.Tab - 1).Enabled = bEnable
      txtVatBk(tabPayment.Tab - 1).Enabled = bEnable
      txtRecharge(tabPayment.Tab - 1).Enabled = bEnable
      
      cmdBkList(tabPayment.Tab - 1).Enabled = bEnable
      cmdTransListBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdUnitListBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdNCBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdDeptBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdProjBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdCCBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdTaxListBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdRechargeBk(tabPayment.Tab - 1).Enabled = bEnable
      cmdUpdateBk(tabPayment.Tab - 1).Enabled = bEnable
End Sub

Private Sub cmdGenAll_Click()
   fraAutoDemandChoice.Left = 1200
   fraAutoDemandChoice.Top = 850
   fraAutoDemandChoice.Visible = True
   fraAutoDemandChoice.ZOrder 0
   fraMain.Enabled = False
End Sub

Private Sub cmdGenerateManual_Click()
   Dim szComp As String, szUnit As String, szBt As String
   Dim szSageAC As String, dtIssueDt As Date, chIC As String
'
   MousePointer = vbHourglass
'
   FlxGridManualConfigure flxAddNewDemands
   If cboTenant.ListCount = 0 Then Call FillcboTenant
   If cboVatCode.ListCount = 0 Then Call FillcboVatCode(cboVatCode)
'
   Set objDemandType = New clsArray
   If iSelectedDemandsRow > 0 Then
'***Add new transaction in previously created demand******
      Call FindLastSelID(szComp, szUnit, szSageAC, dtIssueDt, szBt)
      szCurCompSageAccNum = szSageAC
'
      chIC = InvCre(szLastIDClicked)
      Call FillcboType(cboType, chIC)
      lblDemandType.Caption = IIf(chIC = "I", "Invoice", "Credit")
'
      txtTenantName.text = szSageAC & " / " & szComp
      txtBatchID.text = szBt
      txtIssueDate.text = Format(dtIssueDt, "dd mmmm yyyy")
      txtAddNewSageText.text = "S/L " & szSageAC
      txtDemandID.text = szLastIDClicked
'
      txtUnit.text = szUnit
'*****Fill grid by selected demands****
      Call FillManualChildinGrid(szLastIDClicked, szBt, flxAddNewDemands)
'
      CalSubTotal flxAddNewDemands, txtSubTAmount, txtSubTVAT, txtSubTTotal
'
      bAddNew = False
   '***Load all previous demand type in the object for saving newly
      LoadDemandTypeInObj flxAddNewDemands, cboType
'
      cboTenant.Visible = False
      fraCreateManualDemand.Left = flxDemands.Left
      fraCreateManualDemand.Top = flxDemands.Top
      fraCreateManualDemand.Visible = True
      fraCreateManualDemand.ZOrder 0
      fraDetails.Visible = False
   Else
      fraInvCrChoice.Top = 1320
      fraInvCrChoice.Left = 1400
      fraInvCrChoice.Visible = True
      fraInvCrChoice.ZOrder 0
      fraMain.Enabled = False
   End If
   MousePointer = vbDefault
End Sub

Private Sub cmdManualDmdOk_Click()
   Dim szComp As String, szUnit As String, szBt As String
   Dim szSageAC As String, dtIssueDt As Date
   Dim adoConn As ADODB.Connection

   MousePointer = vbHourglass

   fraInvCrChoice.Visible = False

   Call FillcboType(cboType, IIf(optManualInv.Value Or optManualAdjInv.Value, "I", "C"))
   lblDemandType.Caption = IIf(optManualInv.Value Or optManualAdjInv.Value, "Invoice", "Credit")
   lblDemandType.Caption = IIf(optManualAdjInv.Value Or optManualAdjCrNote.Value, "Adjustment ", "") & lblDemandType.Caption

   bAddNew = True

   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   txtDemandID.text = Format(NextRef(adoConn, "DEMAND_REF"), "00000000")
   txtBatchID.text = NextRef(adoConn, "BATCH_REF")
   flxAddNewDemands.TextMatrix(1, 3) = "M"

   cboTenant.Visible = True
   fraCreateManualDemand.Left = flxDemands.Left - 40
   fraCreateManualDemand.Top = flxDemands.Top
   fraCreateManualDemand.Visible = True
   fraCreateManualDemand.ZOrder 0
   fraDetails.Visible = False

   flxAddNewDemands.TextMatrix(1, 1) = "001"

   adoConn.Close
   Set adoConn = Nothing

   MousePointer = vbDefault
End Sub

Private Sub LoadDemandTypeInObj(conFlxGrid As Control, conCombo As Control)
   Dim iLoop As Integer, iRow As Integer
   Dim szaTemp() As String
'
   For iRow = 1 To conFlxGrid.Rows - 1
      If conFlxGrid.TextMatrix(iRow, 2) <> "" Then
         For iLoop = 0 To conCombo.ListCount - 1
            szaTemp = Split(conCombo.List(iLoop), " / ")
            If conFlxGrid.TextMatrix(iRow, 2) = szaTemp(1) Then Exit For
         Next iLoop
         objDemandType.AddItemPos Trim(szaTemp(1)), iLoop
      End If
   Next iRow
End Sub

Private Sub CalSubTotal(conGrid As Control, conTextTA As Control, conTextVAT As Control, conTextTotal As Control)
   Dim iRow As Integer
'
   conTextTA.text = ""
   conTextVAT.text = ""
   conTextTotal.text = ""
   For iRow = 1 To conGrid.Rows - 1
      conTextTA.text = Format(CCur(IIf(conTextTA.text = "", 0, conTextTA.text)) + CCur(IIf(conGrid.TextMatrix(iRow, 8) = "", 0, conGrid.TextMatrix(iRow, 8))), "0.00")
      conTextVAT.text = Format(CCur(IIf(conTextVAT.text = "", 0, conTextVAT.text)) + CCur(IIf(conGrid.TextMatrix(iRow, 9) = "", 0, conGrid.TextMatrix(iRow, 9))), "0.00")
      conTextTotal.text = Format(CCur(IIf(conTextTotal.text = "", 0, conTextTotal.text)) + CCur(IIf(conGrid.TextMatrix(iRow, 10) = "", 0, conGrid.TextMatrix(iRow, 10))), "0.00")
   Next iRow
End Sub
'
Private Sub FindLastSelID(szComp As String, szUnit As String, szSageAC As String, dtIssueDt As Date, szBt As String)
   Dim adoDBConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
   Dim szSQL As String
'
   Set adoDBConn = New ADODB.Connection
   Set adoRst = New ADODB.Recordset
'
   adoDBConn.Open "DSN=" & Adsn & ";UID=;PWD="
   szSQL = "SELECT DemandRecords.TenantCompanyName as TCN, " & _
               "DemandRecords.UnitNumber as UN, DemandRecords.SageAccountNumber as SAC, " & _
               "DemandRecords.IssueDate as ID, DemandRecords.BatchID as BI " & _
           "FROM DemandRecords, DemandSplitRecords " & _
           "WHERE DEMANDRECORDS.DEMANDID = " & szLastIDClicked & " AND " & _
               "DemandSplitRecords.DemandID = DemandRecords.DemandID;"
   adoRst.Open szSQL, adoDBConn, adOpenDynamic, adLockPessimistic
'
   szComp = adoRst!TCN
   szUnit = adoRst!UN
   szSageAC = adoRst!SAC
   dtIssueDt = adoRst!ID
   szBt = adoRst!BI
'
   adoRst.Close
   adoDBConn.Close
   Set adoRst = Nothing
   Set adoDBConn = Nothing
End Sub
'
Private Sub BatchCombo()
'   Dim adoDBConn As ADODB.Connection
'   Dim adoRstBt As ADODB.Recordset
'   Dim szSql As String
''
'   Set adoDBConn = New ADODB.Connection
'   Set adoRstBt = New ADODB.Recordset
''
'   adoDBConn.Open "DSN=" & Adsn & ";UID=;PWD="
'   szSql = "SELECT BATCHID FROM DemandRecords GROUP BY BATCHID ORDER BY BATCHID DESC"
'   adoRstBt.Open szSql, adoDBConn, adOpenDynamic, adLockPessimistic
''
'   cmbManualBatchNo.AddItem "NONE"
'   If adoRstBt.EOF Then
'      cmbManualBatchNo.AddItem "1"
'   Else
'      While Not adoRstBt.EOF
'         If adoRstBt.Fields(0).Value <> "" Then cmbManualBatchNo.AddItem adoRstBt.Fields(0).Value
'         adoRstBt.MoveNext
'      Wend
'   End If
'   cmbManualBatchNo.AddItem "NEW"
'   cmbManualBatchNo.ListIndex = 1
'   adoRstBt.Close
'   adoDBConn.Close
'   Set adoRstBt = Nothing
'   Set adoDBConn = Nothing
End Sub

Private Sub ComboPosition()
   cboType.Left = flxAddNewDemands.Left + 7360
   cboType.Top = (flxAddNewDemands.RowHeightMin * (flxAddNewDemands.Rows - 1)) + flxAddNewDemands.Top + 10
   cboType.Width = 1240
   cboType.Visible = True
End Sub

Private Sub FillcboVatCode(conCombo As Control)
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String

   conCombo.Clear

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM TLBVATCODE"
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   If rdoRst1.EOF = False Then
       While rdoRst1.EOF = False
           conCombo.AddItem rdoRst1!VAT_ID & " / " & rdoRst1!VAT_CODE & " / " & rdoRst1!VAT_RATE
           rdoRst1.MoveNext
       Wend
   End If

   rdoRst1.Close
   rdoConn.Close
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Sub

Private Sub FillcboType(conCombo As Control, szInvCr As String)
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String

   conCombo.Clear

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT ID, Type " & _
             "FROM DemandTypes " & _
             "WHERE InvCrd = '" & szInvCr & "';"
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
   
   If rdoRst1.EOF = False Then
       While rdoRst1.EOF = False
           conCombo.AddItem rdoRst1!ID & " / " & rdoRst1!Type
           rdoRst1.MoveNext
       Wend
   End If

   rdoRst1.Close
   rdoConn.Close
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Sub

Private Sub cmdMoveFirst_Click()
'move to the first demand in database
'if in print or reprint mode need to update the status of
'sendtoprint for demand currently on screen.
'if in edit mode need to prompt to save changes that have been made
'If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
'If EditMode = True Then Call SaveChanges
'Call EmptyBoxes
'Call GetFirstDemand
End Sub

Private Sub cmdMoveLast_Click()
'move to last demand in database
'if in print or reprint mode need to update the status of
'sendtoprint for demand currently on screen.
'if in edit mode need to prompt to save changes that have been made
'If Text1.text = "" Then
'    MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
'    Exit Sub
''End If
'Dim b As Integer
'b = 1
'
'If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
'If EditMode = True Then Call SaveChanges
'
'Set adoConn = New ADODB.Connection
'Set adoRst = New ADODB.Recordset
''connect to database
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
''get all uniquerefnumbers of demands in program
'strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords"
''if in edit, print or reprint mode get uniquerefnumbers of demands that have correct printed and exported to sage status
'If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UPDATE_SAGE = 'N'"
'If PrintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'N'"
'If ReprintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'Y'"
'
'adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
''find the largest uniquerefnumber and set this to b
'If adoRst.EOF = False Then
'    While adoRst.EOF = False
'        If b < adoRst!UniqueRefNumber Then b = adoRst!UniqueRefNumber
'        adoRst.MoveNext
'    Wend
'End If
'
'adoRst.Close
'
'Set adoRst = Nothing
'Set adoRst = New ADODB.Recordset
''get the demand details for demand with uniquerefnumber b
'strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & b & "';"
'adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
'
'Call EmptyBoxes
'Call GetRecord
'
'adoRst.Close
'adoConn.Close
'Set adoRst = Nothing
'Set adoConn = Nothing

End Sub

Private Sub cmdMoveNext_Click()
   'move to next demand in database
   'if in print or reprint mode need to update the status of
   'sendtoprint for demand currently on screen.
   'if in edit mode need to prompt to save changes that have been made
'   If Text1.text = "" Then
'       MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
'       Exit Sub
'   End If

'   Dim b As Integer
'   Dim last As Boolean
'   last = False
'
'   If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
'   If EditMode = True Then Call SaveChanges
'
'   Set adoConn = New ADODB.Connection
'   Set adoRst = New ADODB.Recordset
'   'connect to database
'   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   'get uniquerefnumbers from demands where uniqueref number is greater than current demand on screen
'   strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber > " & CInt(Text1.text)
   'if in edit, print or reprint mode get uniquerefnumbers where printed and exported to sage are correct status
'   If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber > " & CInt(Text1.text) & " AND UPDATE_SAGE = 'N'"
'   If PrintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'N' AND UniqueRefNumber > " & CInt(Text1.text)
'   If ReprintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'Y' AND UniqueRefNumber > " & CInt(Text1.text)
'
'   adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockPessimistic
'   'work out which is the smallest uniquerefnumber and set to b
'   If adoRst.EOF = False Then
'       adoRst.MoveLast
'       adoRst.MoveFirst
'       If adoRst.RecordCount = 1 Then last = True
'       b = adoRst!UniqueRefNumber
'       While adoRst.EOF = False
'           If b > adoRst!UniqueRefNumber Then b = adoRst!UniqueRefNumber
'           adoRst.MoveNext
'       Wend
'   Else ' if no records so current demand is last demand
'       MsgBox "This is the last demand.", vbOKOnly + vbInformation, "Last Demand"
'       adoRst.Close
'       Set adoRst = Nothing
'       adoConn.Close
'       Set adoConn = Nothing
'       Exit Sub
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'   Set adoRst = New ADODB.Recordset
'   'get details for demand with Uniquerefnumber b
'   strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & b & "';"
'   adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF = True And adoRst.BOF = True Then
'       MsgBox "This is the last demand", vbOKOnly + vbInformation, "Last Demand"
'       adoRst.Close
'       Set adoRst = Nothing
'       adoConn.Close
'       Set adoConn = Nothing
'       Exit Sub
'   End If
'
'   Call EmptyBoxes
'   Call GetRecord
'
'   If last = True Then MsgBox "This is the last demand", vbOKOnly + vbInformation, "Last Demand"
'
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
   
End Sub

Private Sub cmdMovePrevious_Click()

'move to previous demand in database
'if in print or reprint mode need to update the status of
'sendtoprint for demand currently on screen.
'if in edit mode need to prompt to save changes that have been made

'If Text1.text = "" Then
'    MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
'    Exit Sub
'End If
'
'Dim b As Integer
'Dim a As Integer
'a = CInt(Text1.text)
'b = 1
'If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
'If EditMode = True Then Call SaveChanges
'
'Set adoConn = New ADODB.Connection
'Set adoRst = New ADODB.Recordset
''connect to database
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
''get all uniquerefnumbers that are less than that of current demand on screen
'strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber < " & CInt(Text1.text)
''if in edit, print or reprint mode get uniquerefnumbers where printed and exported to sage are correct status
''If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber > " & CInt(Text1.Text) & " AND UPDATE_SAGE = 'N'"
'If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber < " & CInt(Text1.text) & " AND UPDATE_SAGE = 'N' "
'If PrintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'N' AND UniqueRefNumber < " & CInt(Text1.text)
'If ReprintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'Y' AND UniqueRefNumber < " & CInt(Text1.text)
'
'adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
'
''work out which is the biggest of uniquerefnumbers in record set and set to b
'If adoRst.EOF = False Then
'    While adoRst.EOF = False
'        If b < adoRst!UniqueRefNumber Then b = adoRst!UniqueRefNumber
'        adoRst.MoveNext
'    Wend
'Else ' no records so set b current demand ref number
'    b = a
'End If
'
'adoRst.Close
'Set adoRst = Nothing
'
'Set adoRst = New ADODB.Recordset
''get details of demand with ref number b
'strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & b & "';"
'adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
'
'Call EmptyBoxes
'Call GetRecord
'
'If b = a Then MsgBox "This is the first demand", vbOKOnly + vbInformation, "First Demand"
'
'adoRst.Close
'adoConn.Close
'Set adoRst = Nothing
'Set adoConn = Nothing

End Sub

Private Sub cmdPrint_Click()
   fraPrintChoice.Left = flxDemands.Left
   fraPrintChoice.Top = flxDemands.Top
   fraPrintChoice.Visible = True
   Call CmbBatchWiseFill
'   Call PrintDemands
End Sub

Private Sub CmbBatchWiseFill()
'   Dim adoDBConn As ADODB.Connection
'   Dim adoRstBt As ADODB.Recordset
'   Dim szSql As String
'
'   Set adoDBConn = New ADODB.Connection
'   Set adoRstBt = New ADODB.Recordset
'
'   adoDBConn.Open "DSN=" & Adsn & ";UID=;PWD="
'   szSql = "SELECT BATCHID FROM DemandRecords GROUP BY BATCHID ORDER BY BATCHID DESC"
'   adoRstBt.Open szSql, adoDBConn, adOpenDynamic, adLockPessimistic
'
'   If adoRstBt.EOF Then
'      cmbManualBatchNo.AddItem ""
'   Else
'      While Not adoRstBt.EOF
'         If adoRstBt.Fields(0).Value <> "" Then cmbBatchWise.AddItem adoRstBt.Fields(0).Value
'         adoRstBt.MoveNext
'      Wend
'   End If
'   cmbBatchWise.AddItem "NEW"
''   CmbBatchWise.ListIndex = 1
'   adoRstBt.Close
'   adoDBConn.Close
'   Set adoRstBt = Nothing
'   Set adoDBConn = Nothing
End Sub

Public Sub PrintDemands()
'   PrintMode = True
'   Set adoConn = New ADODB.Connection
'   Set adoRst = New ADODB.Recordset
'
'   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'   strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE IsPrinted = 'N'"
'   adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
'   If adoRst.EOF = False Then
'       While adoRst.EOF = False
'           adoRst!SendToPrint = "Y"
'           adoRst.Update
'           adoRst.MoveNext
'       Wend
'   Else
'       MsgBox "There are no demands to print.", vbOKOnly + vbInformation, "Print Demands"
'       adoRst.Close
'       adoConn.Close
'       Set adoRst = Nothing
'       Set adoConn = Nothing
'       PrintMode = False
'       Exit Sub
'   End If
'
'   adoRst.Close
'   adoConn.Close
'   Set adoConn = Nothing
'   Set adoRst = Nothing

'   Call EmptyBoxes
'   Call GetFirstDemand
'
'   'user selected Print Demands
'   lbl1.Visible = True
'   chkPrint.Visible = True
'   chkPrint.Value = 1
'   cmdGenAll.Visible = False
'   cmdGenerateManual.Visible = False
'   cmdEdit.Visible = False
'   cmdDelete.Visible = False
'   cmdDeleteOld.Visible = False
'   cmdReprint.Visible = False
'   cmdPrint.Visible = False
'   cmdPrintThis.Visible = False
'   cmdPrintAll.Visible = True
'   cmdPrintBatch.Visible = False
'   cmdPrintSome.Visible = True
'   cmdCancelPrint.Visible = True
'   Call DisableMenu
End Sub

Private Sub cmdManualDmdCancel_Click()
   fraInvCrChoice.Visible = False
   fraMain.Enabled = True
End Sub

'Private Sub cmdPayee_Click()
'   LoadPayeeInFlxGrid "ALL"
'
'   fraPayee.Left = cmdPayee.Left
'   fraPayee.Top = cmdPayee.Top + cmdPayee.Height - 20
'   fraPayee.Visible = True
'   fraPayee.ZOrder 0
'   txtSearchAC.SetFocus
''   tabDmdRcpt.Enabled = False
'End Sub

'Private Sub ConfigureFlxPayee()
'   flxPayee.Rows = 2
'   flxPayee.Cols = 4
'   flxPayee.Clear
'
'   flxPayee.ColAlignment(0) = vbAlignTop
'   flxPayee.ColWidth(0) = 1500     'Sage A/C
'   flxPayee.ColWidth(1) = 2500     'Tenant Name
'   flxPayee.ColWidth(2) = 1500     'Unit No.
'   flxPayee.ColWidth(3) = 1       'for future
'
'   flxPayee.TextMatrix(0, 0) = "Sage A/C"
'   flxPayee.TextMatrix(0, 1) = "Tenant Name"
'   flxPayee.TextMatrix(0, 2) = "Unit"
'
'   lblPayeeFlxConfigured.Caption = "YES"
''
'   flxPayee.RowHeightMin = 315
'End Sub

'Private Sub LoadPayeeInFlxGrid(szSortingRule As String)
''   If lblPayeeFlxConfigured.Caption = "NOT" Then ConfigureFlxPayee
''
'   Dim szSQL As String
'   Dim adoConTenant As ADODB.Connection
'   Dim adoRstTenant As ADODB.Recordset
''
'   Set adoConTenant = New ADODB.Connection
'   Set adoRstTenant = New ADODB.Recordset
''
''   If lblFlxPayee.Caption <> "EMPTY" Then ConfigureFlxPayee
''
'   If szSortingRule = "ALL" Then
''!!!!    TENANTS.OCCUPIDE  does not exits  !!*********
'      szSQL = "SELECT TENANTS.SageAccountNumber, TENANTS.CompanyName, UNITS.UNITNUMBER " & _
'              "FROM TENANTS, UNITS " & _
'              "WHERE TENANTS.OCCUPIDE = YES AND " & _
'                  "TENANTS.SageAccountNumber = UNITS.SageAccountNumber;"
'   Else
'      szSQL = ""
'   End If
''
'   adoConTenant.Open "DSN=" & Adsn & ";UID=;PWD="
'   adoRstTenant.Open szSQL, adoConTenant, adOpenDynamic, adLockPessimistic
''
'   While Not adoRstTenant.EOF
'      szSQL = ""
'      szSQL = szSQL & adoRstTenant!SageAccountNumber & vbTab & adoRstTenant!CompanyName & vbTab & adoRstTenant!UnitNumber
'      flxPayee.AddItem szSQL, flxPayee.Rows - 1
'      adoRstTenant.MoveNext
'   Wend
'   lblFlxPayee.Caption = "LOADED"
'End Sub

Private Sub cmdNC_Click()
   LoadNominalCodeNC
'
   fraListNC.Left = txtNominalCodeTR.Left + Frame8.Left
   fraListNC.Top = txtNominalCodeTR.Top + txtNominalCodeTR.Height + Frame8.Top
   fraListNC.Visible = True
   fraListNC.ZOrder 0
   flxListNC.SetFocus
End Sub

Private Sub cmdNCBk_Click(Index As Integer)
   LoadNominalCodeBk
'
   fraListBk(tabPayment.Tab - 1).Left = txtNCBk(tabPayment.Tab - 1).Left + fraBkInput(tabPayment.Tab - 1).Left
   fraListBk(tabPayment.Tab - 1).Top = txtNCBk(tabPayment.Tab - 1).Top + txtNCBk(tabPayment.Tab - 1).Height + fraBkInput(tabPayment.Tab - 1).Top
   fraListBk(tabPayment.Tab - 1).Visible = True
   fraListBk(tabPayment.Tab - 1).ZOrder 0
   flxListBk(tabPayment.Tab - 1).SetFocus
   sTextBox = "NC"
End Sub

Private Sub LoadNominalCodeNC()
   flxListNC.ColWidth(0) = 1500
   flxListNC.ColWidth(1) = 2700
   flxListNC.ColAlignment = vbLeftJustify

' Error Handler
  On Error GoTo Error_Handler

  ' Declare Objects
  Dim oSDO As SageDataObject120.SDOEngine
  Dim oWS As SageDataObject120.Workspace
  Dim oNominalRecord As SageDataObject120.NominalRecord

  ' Declare Variables
  Dim szDataPath As String
  Dim bFlag As Boolean

  ' Create the SDO Engine object
  Set oSDO = New SageDataObject120.SDOEngine

  ' Create the workspace
  Set oWS = oSDO.Workspaces.Add("Example")

  ' Select Company.  The SelectCompany method takes the program install
  ' folder as a parameter
  szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
'  szDataPath = oSDO.SelectCompany("C:\Program Files\Sage\Accounts")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If
   
   ' Try to Connect - Will throw an exception if it fails
   If oWS.Connect(szDataPath, "MANAGER", "", "Example") Then
   
     ' Create instance of NominalRecord object
     Set oNominalRecord = oWS.CreateObject("NominalRecord")
   
     ' Call the AddNew method
   '     oNominalRecord.AddNew
   
     oNominalRecord.MoveFirst
     Dim i As Integer
     i = oNominalRecord.Count
     For i = 1 To oNominalRecord.Count
        flxListNC.TextMatrix(i, 0) = CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)
        flxListNC.TextMatrix(i, 1) = CStr(oNominalRecord.Fields.Item("NAME").Value)
        flxListNC.AddItem ""
        oNominalRecord.MoveNext
     Next i
     
     ' Disconnect
     oWS.Disconnect
   End If
   
   ' Destroy Objects
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
   Exit Sub
   
' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text, vbOKOnly, "SDO Examples"
End Sub

Private Sub LoadNominalCodeBk()
   flxListBk(tabPayment.Tab - 1).ColWidth(0) = 1500
   flxListBk(tabPayment.Tab - 1).ColWidth(1) = 2700
   flxListBk(tabPayment.Tab - 1).ColAlignment = vbLeftJustify

' Error Handler
  On Error GoTo Error_Handler

  ' Declare Objects
  Dim oSDO As SageDataObject120.SDOEngine
  Dim oWS As SageDataObject120.Workspace
  Dim oNominalRecord As SageDataObject120.NominalRecord

  ' Declare Variables
  Dim szDataPath As String
  Dim bFlag As Boolean

  ' Create the SDO Engine object
  Set oSDO = New SageDataObject120.SDOEngine

  ' Create the workspace
  Set oWS = oSDO.Workspaces.Add("Example")

  ' Select Company.  The SelectCompany method takes the program install
  ' folder as a parameter
  szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
'  szDataPath = oSDO.SelectCompany("C:\Program Files\Sage\Accounts")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If
   
   ' Try to Connect - Will throw an exception if it fails
   If oWS.Connect(szDataPath, "MANAGER", "", "Example") Then
   
     ' Create instance of NominalRecord object
     Set oNominalRecord = oWS.CreateObject("NominalRecord")
   
     ' Call the AddNew method
   '     oNominalRecord.AddNew
   
     oNominalRecord.MoveFirst
     Dim i As Integer
     i = oNominalRecord.Count
     For i = 1 To oNominalRecord.Count
        flxListBk(tabPayment.Tab - 1).TextMatrix(i, 0) = CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)
        flxListBk(tabPayment.Tab - 1).TextMatrix(i, 1) = CStr(oNominalRecord.Fields.Item("NAME").Value)
        flxListBk(tabPayment.Tab - 1).AddItem ""
        oNominalRecord.MoveNext
     Next i
     
     ' Disconnect
     oWS.Disconnect
   End If
   
   ' Destroy Objects
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
   Exit Sub

' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text, vbOKOnly, "SDO Examples"
End Sub

Private Sub cmdNewBk_Click(Index As Integer)
   If cmdEditBk(tabPayment.Tab - 1).Enabled = False Then Exit Sub
'
'   If MsgBox("Do you want to add new " & IIf(Index = 0, "Invoice", "Credit Note") & "?", vbYesNo) = vbNo Then Exit Sub
   If MsgBox("Do you want to add new Bank Payment?", vbYesNo) = vbNo Then Exit Sub
'
   HandleTextBoxesBk True, True
'
   cmdNewBk(tabPayment.Tab - 1).Enabled = False
   txtBkAc(tabPayment.Tab - 1).SetFocus
   bChangesMade = True
End Sub

Private Sub cmdOKFlxBk_Click(Index As Integer)
   flxListBk_DblClick (tabPayment.Tab - 1)
End Sub

Private Sub cmdOKFlxNC_Click()
   flxListNC_DblClick
End Sub

'Private Sub cmdPayeeClose_Click()
'   fraPayee.Visible = False
''   tabDmdRcpt.Enabled = True
'End Sub

'Private Sub cmdPayeeOK_Click()
''   tabDmdRcpt.Enabled = True
'   flxPay.TextMatrix(flxPay.Row, 1) = flxPayee.TextMatrix(flxPayee.Row, 0)
'   fraPayee.Visible = False
'   cmdPayee.Visible = False
'End Sub

Private Sub cmdPostDemands_Click()
   If MsgBox("Are you sure to post demands?", vbQuestion + vbYesNo, "Demand Posting") = vbNo Then Exit Sub
'
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String
'
   MousePointer = vbHourglass
'
   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt
'
   SQLStr1 = "SELECT COUNT(DemandID) AS P " & _
             "FROM DEMANDRECORDS " & _
             "WHERE DemandRecords.IsPrinted = TRUE AND " & _
                   "DemandRecords.UPDATE_SAGE = TRUE AND " & _
                   "DEMANDHISTORY=FALSE"
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
'
   If rdoRst1!P = 0 Then
      MsgBox "         NO RECORDS TO POST." & (Chr(13) + Chr(10)) & _
             "Demands must be printed and updated to SAGE " & (Chr(13) + Chr(10)) & _
             "before they can be posted!!", vbInformation + vbOKOnly, "POST RECORDS"
      rdoRst1.Close
      rdoConn.Close
'
      Set rdoRst1 = Nothing
      Set rdoConn = Nothing
'
      Exit Sub
   End If
   rdoRst1.Close
'
   SQLStr1 = "UPDATE DemandRecords " & _
             "SET DemandRecords.DemandHistory=TRUE " & _
             "WHERE DemandRecords.IsPrinted = TRUE AND " & _
                  "DemandRecords.UPDATE_SAGE = TRUE"

   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
'
   rdoRst1.Close
   rdoConn.Close
'
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing

   FillDemandsFlxGrid flxDemands, False
'
   MousePointer = vbDefault
   MsgBox "Demands have been posted successfully", vbInformation + vbOKOnly, "Demand Posting"
'
   FlxDemandsConfigure flxDemandHistory
   FillDemandsFlxGrid flxDemandHistory, True    'True - uploading history records
'
   FlxDemandsConfigure flxDemands
   FillDemandsFlxGrid flxDemands, False         'Flase - uploading history, which are already posted and printed and exported to sage
End Sub

Private Sub cmdPrintAll_Click()
'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey
Dim adoConn As ADODB.Connection
Dim adoRst As ADODB.Recordset
Dim strSQLTitles As String

Set adoConn = New ADODB.Connection
Set adoRst = New ADODB.Recordset

adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE IsPrinted = 'N'"
adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic

If adoRst.EOF = False Then
    While adoRst.EOF = False
        'adoRst.EditMode
        adoRst!IsPrinted = "C"
        adoRst!SendToPrint = ""
        adoRst.Update
        adoRst.MoveNext
    Wend
End If
adoRst.Close
adoConn.Close
Set adoRst = Nothing
Set adoConn = Nothing

CR1.ReportFileName = App.Path & "\REPORTFILES\Demand" & SCID & ".rpt"
CR1.printReport

'Call SetPrintedtoYes
'cmdCancelPrint.Visible = False
'cmdPrintAll.Visible = False
'cmdPrintBatch.Visible = True
'cmdPrintSome.Visible = False
'cmdGenAll.Visible = True
'cmdGenerateManual.Visible = True
'cmdEdit.Visible = True
'cmdDelete.Visible = True
'cmdDeleteOld.Visible = True
'cmdPrint.Visible = True
'cmdReprint.Visible = True
'cmdPrintThis.Visible = True
'chkPrint.Visible = False
'lbl1.Visible = False
'Call EnableMenu
'
'PrintMode = False
'
'Call EmptyBoxes
'Call GetFirstDemand
End Sub

Private Sub cmdPrintBatch_Click()
   Call PrintBatchSelected
End Sub

Private Sub cmdPrintSome_Click()
Dim rdoConn As New RDO.rdoConnection
Dim rdoRst1 As rdoResultset
Dim SQLStr1 As String

MousePointer = vbHourglass

Call UpdatePrint

rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
rdoConn.CursorDriver = rdUseIfNeeded
rdoConn.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

If rdoRst1.EOF = False Then
    While rdoRst1.EOF = False
        rdoRst1.Edit
        rdoRst1!IsPrinted = "C"
        rdoRst1!SendToPrint = ""
        rdoRst1.Update
        rdoRst1.MoveNext
    Wend
    rdoRst1.Close
    rdoConn.Close
Else
    MsgBox "There are no demands selected to print!", vbOKOnly + vbInformation, "Print Demands"
    rdoRst1.Close
    rdoConn.Close
    Exit Sub
End If


'*****************Samrat blocked it for temporary, because its not working*******************
'print the selected demands so those that have IsPrinted = 'C'
'CR1.ReportFileName = App.Path & "\REPORTFILES\Demand" & SCID & ".rpt"
'CR1.printReport

'Call SetPrintedtoYes
'cmdCancelPrint.Visible = False
'cmdPrintAll.Visible = False
'cmdPrintBatch.Visible = True
'cmdPrintSome.Visible = False
'cmdGenAll.Visible = True
'cmdGenerateManual.Visible = True
'cmdEdit.Visible = True
'cmdDelete.Visible = True
'cmdDeleteOld.Visible = True
'cmdPrint.Visible = True
'cmdPrintThis.Visible = True
'cmdReprint.Visible = True
'chkPrint.Visible = False
'lbl1.Visible = False
'Call EnableMenu
'
'PrintMode = False
'
'Call EmptyBoxes
'Call GetFirstDemand
'
MousePointer = vbDefault

End Sub

Private Sub cmdPrintThis_Click()
   Dim bPrintMarked As Boolean
   Dim adoConn As ADODB.Connection
   Dim strSQLTitles As String
   Dim varAns
'
   If iSelectedDemandsRow < 1 Then
      MsgBox "Please select atleast one demand from the grid.", vbCritical, "No demand selected"
      Exit Sub
   End If
   bPrintMarked = False
   varAns = MsgBox("Do you want to make the demand as printed?", vbYesNoCancel, "Printed")
   bPrintMarked = IIf(varAns = vbYes, True, False)
   If varAns = vbCancel Then Exit Sub
'
   Dim iRow As Integer, szWhere As String, iTemp As Integer
   Dim adoRstDRecords As ADODB.Recordset, adoRstDRCurrent As ADODB.Recordset
   Dim adoTrans As ADODB.Recordset
'
'Collect & insert in the array all Transaction type
   Dim szaTransactionType() As String
   ReDim szaTransactionType(iTransactionType) As String
'
   Set adoConn = New ADODB.Connection
   Set adoTrans = New ADODB.Recordset
   strSQLTitles = "SELECT * " & _
                  "FROM DemandTypes " & _
                  "WHERE ID > 0;"
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   adoTrans.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'Debug.Print adoTrans!ID
'
   If Not adoTrans.EOF Then
      For iRow = 0 To iTransactionType - 1
         szaTransactionType(CInt(adoTrans!ID)) = adoTrans!Type
         adoTrans.MoveNext
      Next iRow
   End If
   adoTrans.Close
   Set adoTrans = Nothing
'
   Set adoRstDRecords = New ADODB.Recordset
   Set adoRstDRCurrent = New ADODB.Recordset
'
   strSQLTitles = "DELETE * FROM tlbDRCURRENTPRINT;"
   adoRstDRCurrent.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
   strSQLTitles = "SELECT * FROM tlbDRCURRENTPRINT;"
   adoRstDRCurrent.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
   For iRow = 1 To flxDemands.Rows - 1
      If flxDemands.TextMatrix(iRow, 0) = "X" Then
         szWhere = "DemandRecords.DEMANDID = " + CStr(CLng(flxDemands.TextMatrix(iRow, 1))) + ""
'
         strSQLTitles = "SELECT DemandRecords.DEMANDID AS D_ID, DemandRecords.BATCHID AS B_ID, " & _
                           "DemandRecords.SAGEACCOUNTNUMBER AS S_AC, " & _
                           "DemandRecords.UNITNUMBER AS U_NUM, DemandRecords.SOURCE AS SOU, " & _
                           "DemandRecords.TRANSACTIONTYPE AS T_TYPE, " & _
                           "DemandRecords.ISSUEDATE AS IDT, DemandSplitRecords.A_M AS A_M, " & _
                           "DemandSplitRecords.splitid as S_ID, DemandSplitRecords.NOMINALCODEFORAMOUNT AS NCA, " & _
                           "DemandSplitRecords.NOMINALNAMEFORAMOUNT AS NNA, " & _
                           "DemandSplitRecords.NOMINALCODEFORVAT AS NCV, " & _
                           "DemandSplitRecords.NOMINALNAMEFORVAT AS NNV, " & _
                           "DemandSplitRecords.NOMINALCODEFORTOTAL AS NCT, " & _
                           "DemandSplitRecords.NOMINALNAMEFORTOTAL AS NNT, " & _
                           "DemandSplitRecords.AMOUNT AS AMT, DemandSplitRecords.VATAMOUNT AS VAMT, " & _
                           "DemandSplitRecords.TOTALAMOUNT AS TAMT, " & _
                           "DemandSplitRecords.SAGEREF AS SREF, " & _
                           "DemandSplitRecords.DUEDATE AS DDT, " & _
                           "DemandSplitRecords.VATMONTH AS VMTH, " & _
                           "DemandSplitRecords.TYPEOFDEMAND AS TDM, " & _
                           "DemandSplitRecords.DESCRIPTION AS DESCR, " & _
                           "DemandSplitRecords.DEMANDSTATEMENT AS D_ST, " & _
                           "DemandSplitRecords.DATEFROM AS D_FROM, " & _
                           "DemandSplitRecords.DATETO AS D_TO " & _
                        "FROM DemandRecords, DemandSplitRecords " & _
                        "WHERE (" & szWhere & ") AND " & _
                           "DemandSplitRecords.DEMANDID = DemandRecords.DEMANDID " & _
                        "ORDER BY DemandSplitRecords.splitid;"
' Debug.Print strSQLTitles
         adoRstDRecords.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
'      GET N UNDATE THE INVOICE NUMBER
         txtSelectedPrintRef.text = Format(NextRef(adoConn, "PRINT_REF"), "00000000")
         AddNewRef adoConn, "PRINT_REF", CLng(txtSelectedPrintRef.text)
'
         While Not adoRstDRecords.EOF
            adoRstDRCurrent.AddNew
'
            adoRstDRCurrent.Fields("UniqueRefNumber").Value = adoRstDRecords.Fields("D_ID").Value
            adoRstDRCurrent.Fields("Batch").Value = adoRstDRecords.Fields("B_ID").Value
            adoRstDRCurrent.Fields("AutomaticManual").Value = adoRstDRecords.Fields("A_M").Value
            adoRstDRCurrent.Fields("SageAccountNumber").Value = adoRstDRecords.Fields("S_AC").Value
            adoRstDRCurrent.Fields("UnitNumber").Value = adoRstDRecords.Fields("U_NUM").Value
            adoRstDRCurrent.Fields("NominalCodeforAmount").Value = adoRstDRecords.Fields("NCA").Value
            adoRstDRCurrent.Fields("NominalNameforAmount").Value = adoRstDRecords.Fields("NNA").Value
            adoRstDRCurrent.Fields("NominalCodeforVAT").Value = adoRstDRecords.Fields("NCV").Value
            adoRstDRCurrent.Fields("NominalNameforVAT").Value = adoRstDRecords.Fields("NNV").Value
            adoRstDRCurrent.Fields("NominalCodeforTotal").Value = adoRstDRecords.Fields("NCT").Value
            adoRstDRCurrent.Fields("NominalNameforTotal").Value = adoRstDRecords.Fields("NNT").Value
            adoRstDRCurrent.Fields("Source").Value = adoRstDRecords.Fields("SOU").Value
            adoRstDRCurrent.Fields("TransactionType").Value = adoRstDRecords.Fields("T_TYPE").Value
            adoRstDRCurrent.Fields("IssueDate").Value = adoRstDRecords.Fields("IDT").Value
            adoRstDRCurrent.Fields("VATMonth").Value = adoRstDRecords.Fields("VMTH").Value
            adoRstDRCurrent.Fields("Typeofdemand").Value = adoRstDRecords.Fields("TDM").Value
            adoRstDRCurrent.Fields("Reference").Value = adoRstDRecords.Fields("SREF").Value
            adoRstDRCurrent.Fields("Description").Value = adoRstDRecords.Fields("DESCR").Value
            adoRstDRCurrent.Fields("PRINT_ID").Value = txtSelectedPrintRef.text
'
            If adoRstDRecords.Fields("D_ST").Value Then
               adoRstDRCurrent.Fields("DueDate").Value = adoRstDRecords.Fields("DDT").Value
               adoRstDRCurrent.Fields("TotalAmount").Value = adoRstDRecords.Fields("TAMT").Value
               adoRstDRCurrent.Fields("VATAmount").Value = adoRstDRecords.Fields("VAMT").Value
               adoRstDRCurrent.Fields("Amount").Value = adoRstDRecords.Fields("AMT").Value
               adoRstDRCurrent.Fields("DemandFrom").Value = adoRstDRecords.Fields("D_FROM").Value
               adoRstDRCurrent.Fields("DemandTo").Value = adoRstDRecords.Fields("D_TO").Value
            End If
'
            adoRstDRCurrent.Update
'           adoRstDRecords!IsPrinted = IIf(bPrintMarked, "Y", "N")       'chage the printed status in the DemandRecords table
            adoRstDRecords.MoveNext
         Wend
         adoRstDRecords.Close
      End If
   Next iRow
   adoRstDRCurrent.Close
'
   Set adoRstDRCurrent = Nothing
'
   ShowReport App.Path & "\REPORTFILES\InvDemandSngMPage" & CStr(SCID) & ".rpt"
'
   If bPrintMarked Then
      iTemp = iSelectedDemandsRow
      szWhere = ""
      For iRow = 1 To flxDemands.Rows - 1
         If flxDemands.TextMatrix(iRow, 0) = "X" Then
            szWhere = szWhere + "DemandRecords.DEMANDID = " + CStr(CLng(flxDemands.TextMatrix(iRow, 1))) + ""
            iTemp = iTemp - 1
            If iTemp Then szWhere = szWhere + " OR " Else Exit For
         End If
      Next iRow
      strSQLTitles = "UPDATE DemandRecords " & _
                     "SET IsPrinted = TRUE " & _
                     "WHERE " & szWhere
      adoRstDRecords.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
   End If
'
   adoConn.Close
   Set adoRstDRecords = Nothing
   Set adoConn = Nothing
'
   FlxDemandsConfigure flxDemands
   FillDemandsFlxGrid flxDemands, False
   FlxGridManualConfigure flxAddNewDemands
   
End Sub

Private Sub cmdProjBk_Click(Index As Integer)
   If LoadProjBk Then
      fraListBk(tabPayment.Tab - 1).Left = txtProjBk(tabPayment.Tab - 1).Left + fraBkInput(tabPayment.Tab - 1).Left
      fraListBk(tabPayment.Tab - 1).Top = txtProjBk(tabPayment.Tab - 1).Top + txtProjBk(tabPayment.Tab - 1).Height + fraBkInput(tabPayment.Tab - 1).Top
      fraListBk(tabPayment.Tab - 1).Visible = True
      fraListBk(tabPayment.Tab - 1).ZOrder 0
      flxListBk(tabPayment.Tab - 1).SetFocus
      sTextBox = "Proj"
   Else
      txtProjBk(tabPayment.Tab - 1).text = ""
   End If
End Sub

Private Function LoadProjBk() As Boolean
   flxListBk(tabPayment.Tab - 1).ColWidth(0) = 1500
   flxListBk(tabPayment.Tab - 1).ColWidth(1) = 2700
   flxListBk(tabPayment.Tab - 1).ColAlignment = vbLeftJustify
'
'    Error Handler
   On Error GoTo Error_Handler
'
'    Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oProjects As SageDataObject120.Projects
'
'    Declare Variables
   Dim szDataPath As String
'
'    Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine
'
'    Create the Workspace
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
   
      ' Create Objects
      Set oProjects = oWS.CreateProjects
      flxListBk(tabPayment.Tab - 1).Clear
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 0) = "Proj. Ref."
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 1) = "Project Name"
      
      LoadProjBk = True
      If oProjects.Count = 0 Then
         MsgBox "No project code has been created in SAGE", vbCritical, "Project Empty"
         LoadProjBk = False
         GoTo Error_Handler
      End If
      
      Dim rRow As Integer
      For rRow = 1 To oProjects.Count
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 0) = CStr(oProjects.Item(rRow - 1).Reference)
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 1) = CStr(oProjects.Item(rRow - 1).Name)
      Next rRow
   
      'Disconnect
      oWS.Disconnect
   End If

   ' Destroy Objects
   Set oProjects = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Function

   ' Error Handling Code
Error_Handler:
   Set oProjects = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Function

Private Sub cmdRechargeBk_Click(Index As Integer)
   lstYNBk(tabPayment.Tab - 1).Top = 840
   lstYNBk(tabPayment.Tab - 1).Left = 9000
   lstYNBk(tabPayment.Tab - 1).Visible = True
   lstYNBk(tabPayment.Tab - 1).SetFocus
   lstYNBk(tabPayment.Tab - 1).ZOrder 0
End Sub

Private Sub cmdReprint_Click()

Call ReprintDemands

End Sub

Private Sub cmdReprintAll_Click()
Dim adoConn As ADODB.Connection
Dim adoRst As ADODB.Recordset
Dim strSQLTitles As String
'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

Set adoConn = New ADODB.Connection
Set adoRst = New ADODB.Recordset

adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE IsPrinted = 'Y'"
adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic

If adoRst.EOF = False Then
    While adoRst.EOF = False
        'adoRst.EditMode
        adoRst!IsPrinted = "C"
        adoRst!SendToPrint = ""
        adoRst.Update
        adoRst.MoveNext
    Wend
End If
adoRst.Close
adoConn.Close
Set adoRst = Nothing
Set adoConn = Nothing

CR1.ReportFileName = App.Path & "\REPORTFILES\Demand" & SCID & ".rpt"
'CR1.Connect = "DSN=WDLimited01;UID=;PWD=;DBQ=<CRWDC>Database=WDLimited01"
CR1.printReport

Call SetPrintedtoYes
'ReprintMode = False

End Sub

Private Sub cmdReprintSome_Click()
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String
'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

Call UpdatePrint

rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
rdoConn.CursorDriver = rdUseIfNeeded
rdoConn.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

If rdoRst1.EOF = False Then
    While rdoRst1.EOF = False
        rdoRst1.Edit
        rdoRst1!IsPrinted = "C"
        rdoRst1!SendToPrint = ""
        rdoRst1.Update
        rdoRst1.MoveNext
    Wend
    rdoRst1.Close
    rdoConn.Close
Else
    MsgBox "There are no demands selected to print!", vbOKOnly + vbInformation, "Print Demands"
    rdoRst1.Close
    rdoConn.Close
    Exit Sub
End If

CR1.ReportFileName = App.Path & "\REPORTFILES\Demand" & SCID & ".rpt"
CR1.printReport

Call SetPrintedtoYes
End Sub

Private Sub cmdSaveBk_Click(Index As Integer)
   If cmdUpdateBk(tabPayment.Tab - 1).Enabled Then
      MsgBox "Please finish add new or edit data first.", vbInformation + vbOKOnly, "Saving Data"
      Exit Sub
   End If
   If MsgBox("Are you sure to save?", vbYesNo + vbQuestion, "Saving Data") = vbNo Then Exit Sub
'
   Dim szSQL As String
   Dim iRow As Integer
   Dim rdoConn As New RDO.rdoConnection
   Dim Rst1 As rdoResultset
'
   If flxBankPay(tabPayment.Tab - 1).TextMatrix(1, 0) = "" Then
      MsgBox "No data has been saved!"
      Exit Sub
   End If
'
   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt
'
   szSQL = "SELECT * FROM tlbBankPayment"
   Set Rst1 = rdoConn.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
'Add New Records
   Dim lLastID As Long
   For iRow = 1 To flxBankPay(tabPayment.Tab - 1).Rows - 1
      If flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 14) = "" Then
         lLastID = CLng(GetLastID("tlbBankPayment")) + 1
         
         Rst1.AddNew
         Rst1!MY_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
         Rst1!TRAN_ID = CStr(lLastID + (iRow - 1))
         Rst1!BANK_AC = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 0)
         Rst1!TRAN_DATE = Format(flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 1), "DD MMMM YYYY")
         Rst1!TRANS = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 2)
         Rst1!TRAN_TYPE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 3)
         Rst1!UNIT_ID = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 4)
         Rst1!NOMINAL_CODE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 5)
         Rst1!DEPT_ID = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 6)
         Rst1!PROJ_REF = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 7)
         Rst1!COST_CODE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 8)
         Rst1!description = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 9)
         Rst1!NET_AMOUNT = CCur(flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 10))
         Rst1!VAT = Format(CCur(flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 11)), "0.00")
         Rst1!TAX_CODE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 12)
         If flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 13) = "X" Then Rst1!RECHARABLE = True
         Rst1.Update
         
         SetLastID "tlbBankPayment", lLastID
      End If
   Next iRow
   Rst1.Close
   
'Update edited record
   For iRow = 1 To flxBankPay(tabPayment.Tab - 1).Rows - 1
      If flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 16) = "1" Then
         szSQL = "SELECT * " & _
                 "FROM tlbBankPayment " & _
                 "WHERE MY_ID = '" & flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 14) & "';"
         Set Rst1 = rdoConn.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
         Rst1.Edit
         Rst1!MY_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
         Rst1!TRAN_ID = CStr(lLastID + (iRow - 1))
         Rst1!BANK_AC = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 0)
         Rst1!TRAN_DATE = Format(flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 1), "DD MMMM YYYY")
         Rst1!TRANS = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 2)
         Rst1!TRAN_TYPE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 3)
         Rst1!UNIT_ID = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 4)
         Rst1!NOMINAL_CODE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 5)
         Rst1!DEPT_ID = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 6)
         Rst1!PROJ_REF = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 7)
         Rst1!COST_CODE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 8)
         Rst1!description = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 9)
         Rst1!NET_AMOUNT = CCur(flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 10))
         Rst1!TAX_CODE = flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 11)
         Rst1!VAT = Format(CCur(flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 12)), "0.00")
         If flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 13) = "X" Then Rst1!RECHARABLE = True
         Rst1.Update
'
         flxBankPay(tabPayment.Tab - 1).TextMatrix(iRow, 16) = "0"
      End If
   Next iRow
'
   MsgBox "Data has been saved successfully"
   HandleTextBoxesBk True, False
   cmdNewBk(tabPayment.Tab - 1).Enabled = True
'
'   Rst1.Close
   rdoConn.Close
'
   Set Rst1 = Nothing
   Set rdoConn = Nothing
   
   bChangesMade = False
End Sub

Private Sub cmdSaveNew_Click()
   Dim szSage() As String, szSageID As String
   Dim strSQLTitles As String
   Dim iLoop As Integer, iDemandID As Integer
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
'
   If txtTenantName.text = "" Then
      MsgBox "Please select the tenant.", vbCritical + vbOKOnly, "Tenant Selection"
      Exit Sub
   End If
'
   If txtIssueDate.text = "" Then
      MsgBox "Plseas select due date.", vbCritical + vbOKOnly, "Due Date"
      txtDueDate_Click
      Exit Sub
   End If
'
   If flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 4) = "" Then
      MsgBox "Please give the description of the last statement.", vbCritical + vbOKOnly, "Error"
      flxAddNewDemands.Col = 4
      flxAddNewDemands_Click
      Exit Sub
   End If
'
   If Not bAddNew Then
      MsgBox "No new data to save.", vbOKOnly, "Save Msg"
      fraCreateManualDemand.Visible = False
      Exit Sub
   End If
'
   If flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 18) = "" And flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 8) <> "" Then
      MsgBox "Please give the VAT Code.", vbCritical + vbOKOnly, "Error"
      flxAddNewDemands.Col = 18
      flxAddNewDemands_Click
      Exit Sub
   End If
'
   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'**Clear Demand Table according to the deamnd id********************************
   ClearDemand adoConn, CLng(txtDemandID.text)
'
   szSage = Split(txtTenantName.text, " / ")
   szSageID = szSage(0)
'
   Set adoRst = New ADODB.Recordset
'
'Save the header part in the DemandRecords table
'Save all transactions from temp grid to the database
   strSQLTitles = "SELECT * FROM DemandRecords"
   adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
   adoRst.AddNew
   adoRst!DemandId = CLng(txtDemandID.text)
   adoRst!BATCHID = CLng(txtBatchID.text)
   adoRst!SageAccountNumber = szSageID
   adoRst!TenantCompanyName = szSage(1)
   adoRst!UnitNumber = txtUnit.text
   adoRst!TransactionType = CByte(IIf(lblDemandType.Caption = "Invoice", 1, 2))
   adoRst!IssueDate = Format(txtIssueDate.text, "dd mmmm yyyy")
   adoRst!SageText = txtAddNewSageText.text
   adoRst!IsPrinted = False
   If InStr(lblDemandType.Caption, "Adjustment") = 0 Then
      adoRst!UPDATE_SAGE = False
   Else
      adoRst!UPDATE_SAGE = True
   End If
   adoRst!DemandHistory = False
   adoRst.Update
   adoRst.Close
'
'Save the split parts in the split table
   strSQLTitles = "SELECT * FROM DemandSplitRecords"
   adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
   For iLoop = 1 To flxAddNewDemands.Rows - 1
      adoRst.AddNew
      adoRst!SplitID = CInt(flxAddNewDemands.TextMatrix(iLoop, 1))
      adoRst!DemandId = CLng(txtDemandID.text)
      adoRst!A_M = flxAddNewDemands.TextMatrix(iLoop, 3)
      adoRst!NominalCodeforAmount = flxAddNewDemands.TextMatrix(iLoop, 11)
      adoRst!NominalNameforAmount = flxAddNewDemands.TextMatrix(iLoop, 12)
      adoRst!NominalCodeForVAT = flxAddNewDemands.TextMatrix(iLoop, 13)
      adoRst!NominalNameforVAT = flxAddNewDemands.TextMatrix(iLoop, 14)
      adoRst!NominalCodeForTotal = flxAddNewDemands.TextMatrix(iLoop, 15)
      adoRst!NominalNameforTotal = flxAddNewDemands.TextMatrix(iLoop, 16)
      adoRst!Amount = IIf(flxAddNewDemands.TextMatrix(iLoop, 8) = "", Null, flxAddNewDemands.TextMatrix(iLoop, 8))
      adoRst!VATAmount = IIf(flxAddNewDemands.TextMatrix(iLoop, 9) = "", 0, flxAddNewDemands.TextMatrix(iLoop, 9))
      adoRst!TotalAmount = IIf(flxAddNewDemands.TextMatrix(iLoop, 10) = "", 0, flxAddNewDemands.TextMatrix(iLoop, 10))
      adoRst!SageRef = flxAddNewDemands.TextMatrix(iLoop, 17)
      adoRst!DateFrom = Format(flxAddNewDemands.TextMatrix(iLoop, 5), "dd mmmm yyyy")
      adoRst!DateTo = Format(flxAddNewDemands.TextMatrix(iLoop, 6), "dd mmmm yyyy")
      adoRst!DueDate = Format(flxAddNewDemands.TextMatrix(iLoop, 7), "dd mmmm yyyy")
      adoRst!VATMonth = IIf(flxAddNewDemands.TextMatrix(iLoop, 5) <> "", Month(txtIssueDate.text), "")
      adoRst!typeofdemand = CInt(objDemandType.GetItemPos(flxAddNewDemands.TextMatrix(iLoop, 2)))
      adoRst!description = flxAddNewDemands.TextMatrix(iLoop, 4)
      adoRst!DemandStatement = IIf(flxAddNewDemands.TextMatrix(iLoop, 2) = "" And _
                                   flxAddNewDemands.TextMatrix(iLoop, 8) = "", False, True)
      adoRst!VAT_CODE = flxAddNewDemands.TextMatrix(iLoop, 18)
      adoRst!SageDepartment = DepartmentID(szSageID, txtUnit.text, flxAddNewDemands.TextMatrix(iLoop, 2))
      adoRst.Update
   Next iLoop

   adoRst.Close
   Set adoRst = Nothing
   

   MsgBox "Your Manual Demands have been Saved.", vbOKOnly + vbInformation, "Date saving Successful"

   FlxDemandsConfigure flxDemands
   FillDemandsFlxGrid flxDemands, False
   FlxGridManualConfigure flxAddNewDemands
   fraCreateManualDemand.Visible = False
   AddNewRef adoConn, "DEMAND_REF", CLng(txtDemandID.text)   'Update the new demand id in the REF table
   AddNewRef adoConn, "BATCH_REF", CLng(txtBatchID.text)     'Update the new BATCH id in the REF table

   fraMain.Enabled = True
   fraDetails.Visible = True

   Set adoConn = Nothing
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String
'
   conFlxGrid.Cols = 11
   conFlxGrid.Clear
   szHeader$ = "|<ID|<I/C|<Company Name|<Unit No|<Issue Dt.|>Total|<Printed|<SAGE A/C|SAGE|<BATCHID"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 300        'Solid column
   conFlxGrid.ColWidth(1) = 900        'Demand ID
   conFlxGrid.ColWidth(2) = 450        'Transaction Type
   conFlxGrid.ColWidth(3) = 3300       'Comp Name
   conFlxGrid.ColWidth(4) = 1600       'Unit Number
   conFlxGrid.ColWidth(5) = 1000       'Issue Date
   conFlxGrid.ColWidth(6) = 1000       'Total
   conFlxGrid.ColWidth(7) = 800        'Print
   conFlxGrid.ColWidth(8) = 1300       'Sage A/C No.
   conFlxGrid.ColWidth(9) = 800        'Sage
   conFlxGrid.ColWidth(10) = 0          'BATCHID  *This column always at the end for keep BATCHID number
   conFlxGrid.Rows = 2
'
   conFlxGrid.RowHeightMin = 315
End Sub

Private Sub cmdSPayAll_Click()
   Dim iRow As Integer
   
   txtSPaymentTotal.text = "0.00"
   txtChqNo.text = Format(curOpeningBal, "0.00")
   For iRow = 1 To flxSPayment.Rows - 1
      flxSPayment.TextMatrix(iRow, 8) = flxSPayment.TextMatrix(iRow, 7)
      txtSPaymentTotal.text = CCur(txtSPaymentTotal.text) + CCur(flxSPayment.TextMatrix(iRow, 7))
      txtChqNo.text = CCur(txtChqNo.text) + CCur(flxSPayment.TextMatrix(iRow, 7))
   Next iRow
   
   bChangesMade = True
End Sub

Private Sub cmdSPClose_Click()
   Dim X As Integer
   
   If bChangesMade Then
      X = MsgBox("Do you want to save changes?", vbQuestion + vbYesNoCancel, "Data Saving")
      If X = vbCancel Then Exit Sub
      If X = vbYes Then cmdSPSave_Click
   End If
   Unload Me
End Sub

Private Sub cmdSPClose_LostFocus()
   tabPayment.SetFocus
End Sub

Private Sub cmdSPFull_Click()
   If flxSPayment.Row = 0 Then Exit Sub
'
   On Error GoTo ErrorHandler
   flxSPayment.Col = 8
   If flxSPayment.Row <= flxSPayment.Rows - 1 Then
      txtChqNo.text = Format(CCur(txtChqNo.text) - CCur(flxSPayment.TextMatrix(flxSPayment.Row, 8)), "0.00")
      txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) - CCur(flxSPayment.TextMatrix(flxSPayment.Row, 8)), "0.00")
      flxSPayment.TextMatrix(flxSPayment.Row, 8) = flxSPayment.TextMatrix(flxSPayment.Row, 7)
      txtChqNo.text = Format(CCur(txtChqNo.text) + CCur(flxSPayment.TextMatrix(flxSPayment.Row, 8)), "0.00")
      txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + CCur(flxSPayment.TextMatrix(flxSPayment.Row, 8)), "0.00")
'
      flxSPayment.Row = flxSPayment.Row + 1
      flxSPayment_Click
   End If
   Exit Sub
ErrorHandler:
   MsgBox "You have reached the end of the record", vbInformation + vbOKOnly, "Error"
End Sub

Private Sub cmdSPSave_Click()
   If Not bChangesMade Then
      MsgBox "There is no transaction to save.", vbInformation + vbOKOnly, "Save"
      Exit Sub
   End If
   If txtNominalCodeTR.text = "" Then
      MsgBox "Please enter the nominal code for the receipt.", vbCritical + vbOKOnly, "Nominal code missing"
      Exit Sub
   End If
   If MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save") = vbNo Then Exit Sub
'
   Dim iRow As Integer, szSQL As String
   Dim rdoConn As New RDO.rdoConnection
   Dim rstRst As rdoResultset
   Dim szaTenant() As String
   Dim lT_ID As Long
'
'   Dim rdoConn As New RDO.rdoConnection
'   Dim rdoRst1 As rdoResultset
'
   MousePointer = vbHourglass
   
   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt
'
   szaTenant = Split(cmbTenant.text, " \ ")
'
   szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM TLBRECEIPT"
   Set rstRst = rdoConn.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
   lT_ID = CLng(IIf(IsNull(rstRst!TID), 0, rstRst!TID))
   rstRst.Close
'
'  Get the details for the demand type selected
   szSQL = "SELECT * " & _
           "FROM tlbReceipt " & _
           "WHERE SageAccountNumber = '" & szaTenant(0) & "'"
'Debug.Print szSQL
   Set rstRst = rdoConn.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
   For iRow = 1 To flxSPayment.Rows - 1
      If baChangesMade(iRow) Then
         rstRst.AddNew
         rstRst!TransactionID = lT_ID + 1
         rstRst!Type = 3
         rstRst!DemandRef = flxSPayment.TextMatrix(iRow, 11)
         rstRst!SageAccountNumber = flxSPayment.TextMatrix(iRow, 2)
         rstRst!UnitID = flxSPayment.TextMatrix(iRow, 3)
         rstRst!RDate = Format(txtSPDate.text, "dd/mm/yyyy")
         rstRst!DDate = Format(flxSPayment.TextMatrix(iRow, 4), "dd/mm/yyyy")
         rstRst!Ref = flxSPayment.TextMatrix(iRow, 5)
         rstRst!Details = flxSPayment.TextMatrix(iRow, 6)
         rstRst!Amount = CCur(flxSPayment.TextMatrix(iRow, 7))
         rstRst!OSAmount = CCur(flxSPayment.TextMatrix(iRow, 8)) - _
                            CCur(flxSPayment.TextMatrix(iRow, 9)) - _
                            CCur(flxSPayment.TextMatrix(iRow, 10))
         rstRst!ReceiptAmount = CCur(flxSPayment.TextMatrix(iRow, 9))
         rstRst!Discount = CCur(flxSPayment.TextMatrix(iRow, 10))
         rstRst!Allocation = flxSPayment.TextMatrix(iRow, 0)
         rstRst!IsSageUpdate = True
         rstRst!UpDateSage = False
         rstRst!ReceiptView = IIf(rstRst!OSAmount > 0, True, False)
         rstRst!BankCode = Left(cmbBankAc.text, 4)
         rstRst!NominalCode = Left(Trim(txtNominalCodeTR.text), 4)
         rstRst.Update
         lT_ID = lT_ID + 1

         StopViewing flxSPayment.TextMatrix(iRow, 0)
      End If
   Next iRow

   rstRst.Close
   rdoConn.Close
   MousePointer = vbDefault
'
   MsgBox "Data has been saved successfully", vbInformation + vbOKOnly, "Data Saving"
   bChangesMade = False
'
   Set rstRst = Nothing
   Set rdoConn = Nothing
   
   MousePointer = vbHourglass
   SPFlxGridConfigure
   LoadDataInGrid
   txtSPaymentTotal.text = "0.00"
   ReDim baChangesMade(flxSPayment.Rows) As Boolean
   
   MousePointer = vbDefault
End Sub

Private Sub StopViewing(TID As Long)
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim szSQL As String

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt
   
   szSQL = "UPDATE tlbReceipt " & _
           "SET tlbReceipt.ReceiptView = False " & _
           "WHERE tlbReceipt.TransactionID = " & TID & ";"
   Set rdoRst1 = rdoConn.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   
   rdoRst1.Close
   rdoConn.Close
'
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Sub

Private Sub cmdTaxListBk_Click(Index As Integer)
   LoadVATBk

   fraListBk(tabPayment.Tab - 1).Left = txtVatBk(tabPayment.Tab - 1).Left - 400
   fraListBk(tabPayment.Tab - 1).Top = txtVatBk(tabPayment.Tab - 1).Top + txtVatBk(tabPayment.Tab - 1).Height
   fraListBk(tabPayment.Tab - 1).Visible = True
   fraListBk(tabPayment.Tab - 1).ZOrder 0
   flxListBk(tabPayment.Tab - 1).SetFocus
   sTextBox = "VAT"
End Sub

Private Sub LoadVATBk()
   flxListBk(tabPayment.Tab - 1).ColWidth(0) = 1000
   flxListBk(tabPayment.Tab - 1).ColWidth(1) = 2500
   flxListBk(tabPayment.Tab - 1).Width = 3600
   fraListBk(tabPayment.Tab - 1).Width = 3700
'   imgFramListCoseBk.Left = 3400
   cmdOKFlxBk(tabPayment.Tab - 1).Left = 2900
   flxListBk(tabPayment.Tab - 1).TextMatrix(0, 0) = "CODE"
   flxListBk(tabPayment.Tab - 1).TextMatrix(0, 1) = "RATE"
   
   Dim rRow As Integer
   Dim Conn2 As New RDO.rdoConnection
'
   Dim szSQL As String
   Dim rstRec As rdoResultset
'
'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt
'
   szSQL = "SELECT VAT_CODE, VAT_RATE " & _
           "FROM tlbVatCode;"
   Set rstRec = Conn2.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
   If Not rstRec.EOF Then
      flxListBk(tabPayment.Tab - 1).Clear
'
      rstRec.MoveFirst
      flxListBk(tabPayment.Tab - 1).ColAlignment(1) = vbRightJustify
'
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 0) = "VAT Code"
      flxListBk(tabPayment.Tab - 1).TextMatrix(0, 1) = "VAT Rate"
'
      rRow = 1
      While Not rstRec.EOF
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 0) = rstRec!VAT_CODE
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 1) = rstRec!VAT_RATE
         rstRec.MoveNext
         If Not rstRec.EOF Then flxListBk(tabPayment.Tab - 1).AddItem ""
         rRow = rRow + 1
      Wend
   End If
'
   rstRec.Close
   Conn2.Close
   
   Set rstRec = Nothing
   Set Conn2 = Nothing
End Sub

Private Sub cmdTransListBk_Click(Index As Integer)
   lstTypeBk(tabPayment.Tab - 1).Top = txtTypeBk(tabPayment.Tab - 1).Top
   lstTypeBk(tabPayment.Tab - 1).Left = txtTypeBk(tabPayment.Tab - 1).Left
   lstTypeBk(tabPayment.Tab - 1).Visible = True
   lstTypeBk(tabPayment.Tab - 1).SetFocus
   lstTypeBk(tabPayment.Tab - 1).ZOrder 0
End Sub

Private Sub cmdUnitListBk_Click(Index As Integer)
   LoadUnit
'
   fraListBk(tabPayment.Tab - 1).Left = txtUnitBk(tabPayment.Tab - 1).Left + fraBkInput(tabPayment.Tab - 1).Left
   fraListBk(tabPayment.Tab - 1).Top = txtUnitBk(tabPayment.Tab - 1).Top + txtUnitBk(tabPayment.Tab - 1).Height + fraBkInput(tabPayment.Tab - 1).Top
   fraListBk(tabPayment.Tab - 1).Visible = True
   fraListBk(tabPayment.Tab - 1).ZOrder 0
   flxListBk(tabPayment.Tab - 1).SetFocus
   sTextBox = "Unit"
End Sub

Private Sub LoadUnit()
   flxListBk(tabPayment.Tab - 1).TextMatrix(0, 0) = "Unit ID"
   flxListBk(tabPayment.Tab - 1).TextMatrix(0, 1) = "Unit Name"
   flxListBk(tabPayment.Tab - 1).ColWidth(0) = 800
   flxListBk(tabPayment.Tab - 1).ColWidth(1) = 2500
   flxListBk(tabPayment.Tab - 1).ColAlignment(0) = vbRightJustify
   flxListBk(tabPayment.Tab - 1).ColAlignment(1) = vbLeftJustify
   
   Dim rRow As Integer
   Dim Conn2 As New RDO.rdoConnection

   Dim szSQL As String
   Dim rstRec As rdoResultset
   
   'Reset screen to show all the units in cboUnits.
   'Set the RDO Connections to the dataset
   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt
   
   szSQL = "SELECT UnitNumber, UnitName " & _
           "FROM Units " & _
           "ORDER BY UnitNumber"
   Set rstRec = Conn2.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstRec.EOF = False Then
      rRow = 1
      While Not rstRec.EOF
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 0) = rstRec!UnitNumber
         flxListBk(tabPayment.Tab - 1).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!UnitName), "", rstRec!UnitName)
         rstRec.MoveNext
         If Not rstRec.EOF Then flxListBk(tabPayment.Tab - 1).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   Conn2.Close
End Sub

Private Sub cmdUpdateBk_Click(Index As Integer)
   If MsgBox("Do you want to update data?", vbYesNo + vbQuestion, "Update Data") = vbNo Then Exit Sub
'
   If cmdEditBk(tabPayment.Tab - 1).Enabled Then         'Not in Edit mode. New record adding
      If Not (flxBankPay(tabPayment.Tab - 1).Rows = 2 And flxBankPay(tabPayment.Tab - 1).TextMatrix(1, 0) = "") Then
         flxBankPay(tabPayment.Tab - 1).AddItem ""
      End If
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 0) = txtBkAc(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 1) = txtDateBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 2) = txtTypeBk(tabPayment.Tab - 1).text
'      flxBankPay(tabPayment.Tab - 1-1).TextMatrix(flxBankPay(tabPayment.Tab - 1-1).Rows - 1, 3) = IIf(tabPurExp.Tab = 0, "PI", "PC")
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 3) = IIf(tabPayment.Tab - 1 = 0, "BP", "BR")
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 4) = txtUnitBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 5) = txtNCBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 6) = txtDeptBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 7) = txtProjBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 8) = txtCCBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 9) = txtDetailsBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 10) = txtNetBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 11) = txtVatBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 12) = cmdTaxListBk(tabPayment.Tab - 1).Caption
      flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Rows - 1, 13) = IIf(txtRecharge(tabPayment.Tab - 1).text = "NO", "", "X")
      
      If MsgBox("Do you want to Add more new data?", vbYesNo + vbQuestion, "Add new") = vbYes Then
         HandleTextBoxesBk True, True
      Else
         HandleTextBoxesBk True, False
         cmdSaveBk(tabPayment.Tab - 1).SetFocus
      End If
   Else
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 0) = txtBkAc(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 1) = txtDateBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 2) = txtTypeBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 3) = "BP"
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 4) = txtUnitBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 5) = txtNCBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 6) = txtDeptBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 7) = txtProjBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 8) = txtCCBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 9) = txtDetailsBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 10) = txtNetBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 11) = cmdTaxListBk(tabPayment.Tab - 1).Caption
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 12) = txtVatBk(tabPayment.Tab - 1).text
      flxBankPay(tabPayment.Tab - 1).TextMatrix(iCurEditRow, 13) = IIf(txtRecharge(tabPayment.Tab - 1).text = "NO", "", "X")
      
      HandleTextBoxesBk True, False
      cmdEditBk(tabPayment.Tab - 1).Enabled = True
   End If
   FlxSumUp flxBankPay(tabPayment.Tab - 1), 10, 12, txtBkTotalNet(tabPayment.Tab - 1), txtBkTotalVat(tabPayment.Tab - 1)
End Sub

Private Sub FlxSumUp(conFlxGrid As Control, iNet As Integer, iVat As Integer, conTextBoxNet As Control, conTextBoxVat As Control)
   Dim iRow As Integer, dNet As Double, dVat As Double
'
   dNet = 0
   dVat = 0
'
   For iRow = 1 To conFlxGrid.Rows - 1
      dNet = dNet + Val(conFlxGrid.TextMatrix(iRow, iNet))
      dVat = dVat + Val(conFlxGrid.TextMatrix(iRow, iVat))
   Next iRow
'
   conTextBoxNet.text = CStr(Format(dNet, "0.00"))
   conTextBoxVat.text = CStr(Format(dVat, "0.00"))
End Sub

Private Sub cmdUpdateDemand_Click()
   Dim szSage() As String, szSageID As String
   Dim strSQLTitles As String
   Dim iLoop As Integer, iDemandID As Integer
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
'
   If txtEditIssueDate.text = "" Then
      MsgBox "Plseas select Issue date.", vbCritical + vbOKOnly, "Due Date"
      txtDueDate_Click
      Exit Sub
   End If
'
   If flxEditDemand.TextMatrix(flxEditDemand.Rows - 1, 18) = "" And flxEditDemand.TextMatrix(flxEditDemand.Rows - 1, 8) <> "" Then
      MsgBox "Please give the VAT Code.", vbCritical + vbOKOnly, "Error"
      flxEditDemand.Col = 18
      flxEditDemand_Click
      Exit Sub
   End If
'
   Set objDemandType = New clsArray
'***Load all previous demand type in the object for saving newly
   LoadDemandTypeInObj flxEditDemand, cboEditType

   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'****Clear Demand Table according to the deamnd id********************************
   ClearDemand adoConn, CLng(txtEditDemandID.text)
'
   szSage = Split(txtEditTenantName.text, " / ")
   szSageID = szSage(0)
'
   Set adoRst = New ADODB.Recordset

'Save the header part in the DemandRecords table
'Save all transactions from temp grid to the database
   strSQLTitles = "SELECT * FROM DemandRecords"
   adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
   adoRst.AddNew
   adoRst!DemandId = CLng(txtEditDemandID.text)
   adoRst!BATCHID = CLng(txtEditBatch.text)
   adoRst!SageAccountNumber = szSageID
   adoRst!TenantCompanyName = szSage(1)
   adoRst!UnitNumber = txtEditUnit.text
'   adoRst!Source=1                     dont delete now, be sure why do we need this field
   adoRst!TransactionType = IIf(Right(lblTransactionType.Caption, 3) = "INV", 1, 2)
   adoRst!IssueDate = Format(txtEditIssueDate.text, "dd mmmm yyyy")
   adoRst!SageText = txtEditAddNewSageText.text
   adoRst!IsPrinted = False
'   adoRst!SendToPrint = False         dont delete now, be sure why do we need this field
   adoRst!UPDATE_SAGE = False
   adoRst!DemandHistory = False
'   adoRst!ExportedToExcel = False     dont delete now, be sure do we need this field
   adoRst.Update
   adoRst.Close
'
'Save the split parts in the split table
   strSQLTitles = "SELECT * FROM DemandSplitRecords"
   adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
   For iLoop = 1 To flxEditDemand.Rows - 1
      adoRst.AddNew
      adoRst!SplitID = CInt(flxEditDemand.TextMatrix(iLoop, 1))
      adoRst!DemandId = CLng(txtEditDemandID.text)
      adoRst!A_M = flxEditDemand.TextMatrix(iLoop, 3)
      adoRst!NominalCodeforAmount = flxEditDemand.TextMatrix(iLoop, 11)
      adoRst!NominalNameforAmount = flxEditDemand.TextMatrix(iLoop, 12)
      adoRst!NominalCodeForVAT = flxEditDemand.TextMatrix(iLoop, 13)
      adoRst!NominalNameforVAT = flxEditDemand.TextMatrix(iLoop, 14)
      adoRst!NominalCodeForTotal = flxEditDemand.TextMatrix(iLoop, 15)
      adoRst!NominalNameforTotal = flxEditDemand.TextMatrix(iLoop, 16)
      adoRst!Amount = IIf(flxEditDemand.TextMatrix(iLoop, 8) = "", Null, flxEditDemand.TextMatrix(iLoop, 8))
      adoRst!VATAmount = IIf(flxEditDemand.TextMatrix(iLoop, 9) = "", 0, flxEditDemand.TextMatrix(iLoop, 9))
      adoRst!TotalAmount = IIf(flxEditDemand.TextMatrix(iLoop, 10) = "", 0, flxEditDemand.TextMatrix(iLoop, 10))
      adoRst!SageRef = flxEditDemand.TextMatrix(iLoop, 17)
      adoRst!DateFrom = Format(flxEditDemand.TextMatrix(iLoop, 5), "dd mmmm yyyy")
      adoRst!DateTo = Format(flxEditDemand.TextMatrix(iLoop, 6), "dd mmmm yyyy")
      adoRst!DueDate = Format(flxEditDemand.TextMatrix(iLoop, 7), "dd mmmm yyyy")
      adoRst!VATMonth = IIf(flxEditDemand.TextMatrix(iLoop, 5) <> "", Month(txtEditIssueDate.text), "")
      adoRst!typeofdemand = CInt(objDemandType.GetItemPos(flxEditDemand.TextMatrix(iLoop, 2)))
      adoRst!description = flxEditDemand.TextMatrix(iLoop, 4)
      adoRst!DemandStatement = IIf(flxEditDemand.TextMatrix(iLoop, 2) = "" And _
                                   flxEditDemand.TextMatrix(iLoop, 8) = "", False, True)
      adoRst!VAT_CODE = flxEditDemand.TextMatrix(iLoop, 18)
      adoRst.Update
   Next iLoop
'
   adoRst.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
'
   MsgBox "Demands have been updated successfully.", vbOKOnly + vbInformation, "Date saving Successful"
'
   FlxDemandsConfigure flxDemands
   FillDemandsFlxGrid flxDemands, False
   FlxGridManualConfigure flxEditDemand
'   fraCreateManualDemand.Visible = False
'
'   AddNewRef adoConn,  "DEMAND_REF", CLng(txtEditDemandID.text)  'Update the new demand id in the REF table
'   AddNewRef adoConn,  "BATCH_REF", CLng(txtEditBatch.text)    'Update the new BATCH id in the REF table
'********************   FIXED CODE        ********************
   fraEditDemandWindow.Visible = False
   cmdEdit.Enabled = True
   fraDetails.Visible = True
End Sub

Private Sub dtDueDate_DateClick(ByVal DateClicked As Date)
   txtIssueDate.text = Format(DateClicked, "dd mmmm yyyy")
End Sub

Private Sub dtEditDate_DateClick(ByVal DateClicked As Date)
   flxEditDemand.TextMatrix(flxEditDemand.Row, flxEditDemand.Col) = Format(DateClicked, "dd/mm/yy")
   dtEditDate.Visible = False
End Sub

Private Sub dtEditDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then dtEditDate.Visible = False
End Sub

Private Sub dtEditDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then dtEditDate.Visible = False
End Sub

Private Sub dtEditDate_LostFocus()
   dtEditDate.Visible = False
End Sub

Private Sub dtEditIssueDate_DateClick(ByVal DateClicked As Date)
   txtEditIssueDate.text = Format(DateClicked, "dd mmmm yyyy")
   dtEditIssueDate.Visible = False
End Sub

Private Sub dtEditIssueDate_GotFocus()
   dtEditIssueDate.Value = CDate(Date)
End Sub

Private Sub dtEditIssueDate_LostFocus()
   dtEditIssueDate.Visible = False
End Sub

Private Sub dtSPDate_DateClick(ByVal DateClicked As Date)
   txtSPDate.text = Format(dtSPDate.Value, "dd/mm/yyyy")
   dtSPDate.Visible = False
End Sub

Private Sub flxAddNewDemands_Click()
   Dim iRow, iLeft, iCol As Integer
'
'  Only user can access the last row of the grid
   If flxAddNewDemands.Row <> flxAddNewDemands.Rows - 1 Then Exit Sub
   If Not bAddNew Then Exit Sub
'
   iLeft = 0
   For iCol = 0 To flxAddNewDemands.ColSel - 1
      iLeft = iLeft + flxAddNewDemands.ColWidth(iCol)
   Next iCol
   Select Case flxAddNewDemands.ColSel
      Case 2:
         If txtIssueDate.text = "" Then
            MsgBox "Please provide Issue Date.", vbInformation + vbOKOnly, "Issue Date"
            Exit Sub
         End If
         cboType.Left = iLeft + flxAddNewDemands.Left + 20
         cboType.Top = flxAddNewDemands.Top + (flxAddNewDemands.RowHeight(flxAddNewDemands.Row) * _
                                 flxAddNewDemands.Row) + 20
         cboType.Width = flxAddNewDemands.ColWidth(flxAddNewDemands.ColSel)
         cboType.Visible = True
         cboType.ZOrder 0
         cboType.SetFocus

      Case 4:                   'Description
         txtAddNewDescription.Left = iLeft + flxAddNewDemands.Left + 40
         txtAddNewDescription.Top = flxAddNewDemands.Top + (flxAddNewDemands.RowHeight(flxAddNewDemands.Row) * _
                                          flxAddNewDemands.Row) + 40
         txtAddNewDescription.Width = flxAddNewDemands.ColWidth(flxAddNewDemands.ColSel) - 40
         txtAddNewDescription.Height = flxAddNewDemands.RowHeightMin * 2
         txtAddNewDescription.Alignment = vbLeftJustify
         txtAddNewDescription.text = flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel)
         txtAddNewDescription.Visible = True
         txtAddNewDescription.ZOrder 0
         txtAddNewDescription.SetFocus
'
      Case 5:           'From Date
         txtDate.Left = iLeft + flxAddNewDemands.Left + 40
         txtDate.Top = flxAddNewDemands.Top + (flxAddNewDemands.RowHeight(flxAddNewDemands.Row) * _
                              flxAddNewDemands.Row) + 60
         txtDate.Width = flxAddNewDemands.ColWidth(flxAddNewDemands.ColSel) - 40
         txtDate.text = flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel)
         txtDate.Visible = True
         txtDate.SetFocus
         txtDate.ZOrder 0
'
      Case 6:           'To Date
         txtDate.Left = iLeft + flxAddNewDemands.Left + 40
         txtDate.Top = flxAddNewDemands.Top + (flxAddNewDemands.RowHeight(flxAddNewDemands.Row) * _
                              flxAddNewDemands.Row) + 60
         txtDate.Width = flxAddNewDemands.ColWidth(flxAddNewDemands.ColSel) - 40
         txtDate.text = flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel)
         txtDate.Visible = True
         txtDate.SetFocus
         txtDate.ZOrder 0
'
      Case 7:           'Due Date
         txtDate.Left = iLeft + flxAddNewDemands.Left + 40
         txtDate.Top = flxAddNewDemands.Top + (flxAddNewDemands.RowHeight(flxAddNewDemands.Row) * _
                              flxAddNewDemands.Row) + 60
         txtDate.Width = flxAddNewDemands.ColWidth(flxAddNewDemands.ColSel) - 40
         txtDate.text = flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel)
         txtDate.Visible = True
         txtDate.SetFocus
         txtDate.ZOrder 0
'
      Case 8:                          'Amount text box
         If flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, 2) = "" Then
            flxAddNewDemands.Col = 2
            flxAddNewDemands_Click
            Exit Sub
         End If
         txtAddNewAmount.Left = iLeft + flxAddNewDemands.Left + 40
         txtAddNewAmount.Top = flxAddNewDemands.Top + (flxAddNewDemands.RowHeight(flxAddNewDemands.Row) * _
                                 flxAddNewDemands.Row) + 40
         txtAddNewAmount.Width = flxAddNewDemands.ColWidth(flxAddNewDemands.ColSel) - 40
         txtAddNewAmount.Alignment = vbRightJustify
         txtAddNewAmount.text = flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel)
         txtAddNewAmount.Visible = True
         txtAddNewAmount.ZOrder 0
         txtAddNewAmount.SetFocus
'
      Case 18:                            'VAT Code
         cboVatCode.Left = iLeft + flxAddNewDemands.Left + 20 '- flxAddNewDemands.ColWidth(1)
         cboVatCode.Top = flxAddNewDemands.Top + (flxAddNewDemands.RowHeight(flxAddNewDemands.Row) * flxAddNewDemands.Row) + 20
         cboVatCode.Width = flxAddNewDemands.ColWidth(flxAddNewDemands.ColSel)
         cboVatCode.Visible = True
         cboVatCode.ZOrder 0
         cboVatCode.SetFocus
'
   End Select
End Sub

Private Sub BankFlxGridConfigure(conFlxGrid As Control)
   Dim szHeader As String
   
   szHeader$ = "<Bank |<Date |<Type |<Trans |<Unit ID |<N/C |<Dept |<Proj Ref |<Cost Code |<Details |>Net |<T/C |>TAX |<RC"

   conFlxGrid.Clear
   conFlxGrid.Cols = 17
'
   conFlxGrid.FormatString = szHeader$
'
   conFlxGrid.ColWidth(0) = 1000             '"Bank"
   conFlxGrid.ColWidth(1) = 1000             '"Date"
   conFlxGrid.ColWidth(2) = 900              '"Type"
   conFlxGrid.ColWidth(3) = 600              '"Trans"
   conFlxGrid.ColWidth(4) = 900              '"Unit ID"
   conFlxGrid.ColWidth(5) = 800              '"N/C"
   conFlxGrid.ColWidth(6) = 800              '"Dept"
   conFlxGrid.ColWidth(7) = 1000             '"Proj Ref"
   conFlxGrid.ColWidth(8) = 1000             '"Cost Code"
   conFlxGrid.ColWidth(9) = 1400             '"Details"
   conFlxGrid.ColWidth(10) = 800             '"Net"
   conFlxGrid.ColWidth(11) = 600             '"T/C"
   conFlxGrid.ColWidth(12) = 800             '"TAX"
   conFlxGrid.ColWidth(13) = 400             '"RC"       'Rechargeable
'
   conFlxGrid.ColWidth(14) = 1           'Stores BkPay Id hidenly
   conFlxGrid.ColWidth(15) = 1           'Marked X when row will be selected
   conFlxGrid.ColWidth(16) = 1           'keep value 0 or 1 for edit, 1=edit
'
   conFlxGrid.RowHeight(1) = 285
'
   conFlxGrid.RowHeightMin = 300
End Sub

Private Sub flxBankPay_DblClick(Index As Integer)
   If cmdEditBk(tabPayment.Tab - 1).Enabled = True Then cmdEditBk_Click (tabPayment.Tab - 1)
End Sub

Private Sub flxBankPay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If cmdEditBk(tabPayment.Tab - 1).Enabled = False Then Exit Sub     'THE GRID IN THE EDIT MODE
   
   If cmdNewBk(tabPayment.Tab - 1).Enabled And (iSelected = 0 Or flxBankPay(tabPayment.Tab - 1).TextMatrix(flxBankPay(tabPayment.Tab - 1).Row, flxBankPay(tabPayment.Tab - 1).Cols - 2) = "X") Then
      iSelected = iSelected + Select1RowFlxGrid(flxBankPay(tabPayment.Tab - 1), flxBankPay(tabPayment.Tab - 1).Row, flxBankPay(tabPayment.Tab - 1).Cols - 2)
      Exit Sub
   End If
   If cmdNewBk(tabPayment.Tab - 1).Enabled And iSelected = 1 Then
      MsgBox "You can edit only ONE data at a time.", vbInformation + vbOKOnly, "Edit Record"
      Exit Sub
   End If
   
   If flxBankPay(tabPayment.Tab - 1).Row <> flxBankPay(tabPayment.Tab - 1).Rows - 1 Then Exit Sub
End Sub

Private Sub flxDemandHistory_Click()
   Call SelectFlxGridRow(flxDemandHistory, flxDemandHistory.RowSel)
   FlxGridConfigure flxChildDemandHistory
'
   Call FillChildinGrid(flxDemandHistory.TextMatrix(flxDemandHistory.RowSel, 1), flxChildDemandHistory)
End Sub

Private Function SelectedID() As String
   Dim iRow As Integer
   
   SelectedID = ""
   For iRow = 1 To flxDemands.Rows - 1
      If flxDemands.TextMatrix(iRow, 0) = "X" Then
         SelectedID = flxDemands.TextMatrix(iRow, 1)
         Exit For
      End If
   Next iRow
End Function

Private Sub FillChildinGrid(szPID As String, conFlxGrid As Control)
   If szPID = "" Then Exit Sub
'
   Dim szStr As String, iRow As Integer
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
'
   Set adoConn = New ADODB.Connection
   Set adoRst = New ADODB.Recordset
'
'   connect to database
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'***get sendtoprint field for all demands that were set to be sent to print***
   szStr = "SELECT DEMANDSPLITRECORDS.SPLITID AS S_ID, " & _
               "DEMANDSPLITRECORDS.A_M AS A_M, " & _
               "DEMANDSPLITRECORDS.NOMINALCODEFORAMOUNT AS NC_A, " & _
               "DEMANDSPLITRECORDS.NOMINALNAMEFORAMOUNT AS NN_A, " & _
               "DEMANDSPLITRECORDS.NOMINALCODEFORVAT AS NC_V, " & _
               "DEMANDSPLITRECORDS.NOMINALNAMEFORVAT AS NN_V, " & _
               "DEMANDSPLITRECORDS.NOMINALCODEFORTOTAL AS NC_T, " & _
               "DEMANDSPLITRECORDS.NOMINALNAMEFORTOTAL AS NN_T, " & _
               "DEMANDSPLITRECORDS.AMOUNT AS AMT, " & _
               "DEMANDSPLITRECORDS.VATAMOUNT AS VAMT, " & _
               "DEMANDSPLITRECORDS.TOTALAMOUNT AS TAMT, " & _
               "DEMANDSPLITRECORDS.DUEDATE AS D_DATE, " & _
               "DEMANDSPLITRECORDS.DESCRIPTION AS DESCP, " & _
               "DEMANDSPLITRECORDS.SAGEREF AS SREF, " & _
               "DEMANDSPLITRECORDS.DATEFROM AS D_FROM, " & _
               "DEMANDSPLITRECORDS.DATETO AS D_TO " & _
           "FROM DEMANDRECORDS, DEMANDSPLITRECORDS " & _
           "WHERE DEMANDRECORDS.DEMANDID = " & CLng(szPID) & " AND " & _
               "DEMANDRECORDS.DEMANDID = DEMANDSPLITRECORDS.DEMANDID " & _
           "ORDER BY DEMANDSPLITRECORDS.SPLITID;"

   adoRst.Open szStr, adoConn, adOpenDynamic, adLockPessimistic
'
'   If Not adoRst.EOF Then
   iRow = 1
   While Not adoRst.EOF
      conFlxGrid.TextMatrix(iRow, 1) = Format(adoRst!S_ID, "000")
      conFlxGrid.TextMatrix(iRow, 2) = adoRst!A_M
      conFlxGrid.TextMatrix(iRow, 3) = adoRst!DESCP
      conFlxGrid.TextMatrix(iRow, 4) = IIf(IsNull(adoRst!D_FROM), "", adoRst!D_FROM)
      conFlxGrid.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!D_TO), "", adoRst!D_TO)
      conFlxGrid.TextMatrix(iRow, 6) = IIf(IsNull(adoRst!D_DATE), "", adoRst!D_DATE)
      conFlxGrid.TextMatrix(iRow, 7) = IIf(IsNull(adoRst!AMT), "", adoRst!AMT)
      conFlxGrid.TextMatrix(iRow, 8) = IIf(adoRst!VAMT = 0, "", adoRst!VAMT)
      conFlxGrid.TextMatrix(iRow, 9) = IIf(adoRst!TAMT = 0, "", adoRst!TAMT)
      conFlxGrid.TextMatrix(iRow, 10) = IIf(IsNull(adoRst!NC_A), "", adoRst!NC_A)
      conFlxGrid.TextMatrix(iRow, 11) = IIf(IsNull(adoRst!NN_A), "", adoRst!NN_A)
      conFlxGrid.TextMatrix(iRow, 12) = IIf(IsNull(adoRst!NC_V), "", adoRst!NC_V)
      conFlxGrid.TextMatrix(iRow, 13) = IIf(IsNull(adoRst!NN_V), "", adoRst!NN_V)
      conFlxGrid.TextMatrix(iRow, 14) = IIf(IsNull(adoRst!NC_T), "", adoRst!NC_T)
      conFlxGrid.TextMatrix(iRow, 15) = IIf(IsNull(adoRst!NN_T), "", adoRst!NN_T)
'      conFlxGrid.TextMatrix(iRow, 14) = szPID
      conFlxGrid.TextMatrix(iRow, 17) = IIf(IsNull(adoRst!SREF), "", adoRst!SREF)          'SageRef FROM SPLIT TABLE
'
      adoRst.MoveNext
      If Not adoRst.EOF Then conFlxGrid.AddItem ""
      iRow = iRow + 1
   Wend
   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub FillManualChildinGrid(szDemandID As String, szBt As String, conFlxGrid As Control)
   Dim szStr As String, iRow As Integer
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
'
   Set adoConn = New ADODB.Connection
   Set adoRst = New ADODB.Recordset
'
   'connect to database
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   'get sendtoprint field for all demands that were set to be sent to print
   szStr = "SELECT DEMANDSPLITRECORDS.SPLITID AS S_ID, " & _
               "DEMANDSPLITRECORDS.A_M AS A_M, " & _
               "DEMANDSPLITRECORDS.NOMINALCODEFORAMOUNT AS NC_A, " & _
               "DEMANDSPLITRECORDS.NOMINALNAMEFORAMOUNT AS NN_A, " & _
               "DEMANDSPLITRECORDS.NOMINALCODEFORVAT AS NC_V, " & _
               "DEMANDSPLITRECORDS.NOMINALNAMEFORVAT AS NN_V, " & _
               "DEMANDSPLITRECORDS.NOMINALCODEFORTOTAL AS NC_T, " & _
               "DEMANDSPLITRECORDS.NOMINALNAMEFORTOTAL AS NN_T, " & _
               "DEMANDSPLITRECORDS.AMOUNT AS AMT, " & _
               "DEMANDSPLITRECORDS.VATAMOUNT AS VAMT, " & _
               "DEMANDSPLITRECORDS.TOTALAMOUNT AS TAMT, " & _
               "DEMANDSPLITRECORDS.DUEDATE AS D_DATE, " & _
               "DEMANDSPLITRECORDS.DESCRIPTION AS DESCP, " & _
               "DEMANDSPLITRECORDS.SAGEREF AS SREF, " & _
               "DEMANDSPLITRECORDS.VAT_CODE AS V_CODE, " & _
               "DEMANDSPLITRECORDS.DATEFROM AS D_FROM, DEMANDSPLITRECORDS.DATETO AS D_TO, " & _
               "DEMANDTYPES.TYPE AS TP, " & _
               "DEMANDRECORDS.BATCHID AS BT, " & _
               "DEMANDRECORDS.SAGETEXT AS ST " & _
           "FROM DEMANDRECORDS, DEMANDSPLITRECORDS, DEMANDTYPES " & _
           "WHERE DEMANDRECORDS.DEMANDID = " & CLng(szDemandID) & " AND " & _
               "DEMANDRECORDS.BATCHID = " & CLng(szBt) & " AND " & _
               "DEMANDRECORDS.DEMANDID = DEMANDSPLITRECORDS.DEMANDID " & _
               "AND DEMANDSPLITRECORDS.TYPEOFDEMAND=DEMANDTYPES.ID " & _
           "ORDER BY DEMANDSPLITRECORDS.SPLITID;"
'Debug.Print szStr
   adoRst.Open szStr, adoConn, adOpenDynamic, adLockPessimistic
'
   iRow = 1
   While Not adoRst.EOF
      conFlxGrid.TextMatrix(iRow, 1) = adoRst!S_ID
      conFlxGrid.TextMatrix(iRow, 2) = IIf(IsNull(adoRst!TP), "", adoRst!TP)
      conFlxGrid.TextMatrix(iRow, 3) = adoRst!A_M
      conFlxGrid.TextMatrix(iRow, 4) = adoRst!DESCP
      conFlxGrid.TextMatrix(iRow, 5) = Format(adoRst!D_FROM, "dd/mm/yyyy")
      conFlxGrid.TextMatrix(iRow, 6) = Format(adoRst!D_TO, "dd/mm/yyyy")
      conFlxGrid.TextMatrix(iRow, 7) = Format(adoRst!D_DATE, "dd/mm/yyyy")
      conFlxGrid.TextMatrix(iRow, 8) = IIf(IsNull(adoRst!AMT), "", adoRst!AMT)
      conFlxGrid.TextMatrix(iRow, 9) = IIf(adoRst!VAMT = 0, "", adoRst!VAMT)
      conFlxGrid.TextMatrix(iRow, 10) = IIf(adoRst!TAMT = 0, "", adoRst!TAMT)
'Other Hidden data
      conFlxGrid.TextMatrix(iRow, 11) = IIf(IsNull(adoRst!NC_A), "", adoRst!NC_A)
      conFlxGrid.TextMatrix(iRow, 12) = IIf(IsNull(adoRst!NN_A), "", adoRst!NN_A)
      conFlxGrid.TextMatrix(iRow, 13) = IIf(IsNull(adoRst!NC_V), "", adoRst!NC_V)
      conFlxGrid.TextMatrix(iRow, 14) = IIf(IsNull(adoRst!NN_V), "", adoRst!NN_V)
      conFlxGrid.TextMatrix(iRow, 15) = IIf(IsNull(adoRst!NC_T), "", adoRst!NC_T)
      conFlxGrid.TextMatrix(iRow, 16) = IIf(IsNull(adoRst!NN_T), "", adoRst!NN_T)
      conFlxGrid.TextMatrix(iRow, 17) = IIf(IsNull(adoRst!SREF), "", adoRst!SREF)
      conFlxGrid.TextMatrix(iRow, 18) = IIf(IsNull(adoRst!V_CODE), "", adoRst!V_CODE)
'
      adoRst.MoveNext
      If Not adoRst.EOF Then conFlxGrid.AddItem ""
      iRow = iRow + 1
   Wend
   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub flxDemands_RowColChange()
   Dim iCol As Integer, iKol As Integer
'
   If flxDemands.TextMatrix(flxDemands.Row, 1) = "" Then Exit Sub
   iIncDec = SelectFlxGridRow(flxDemands, flxDemands.Row) 'Returns 1 or -1 depends on selection
   If iIncDec < 1 And chkSelectAllDemands.Value Then chkSelectAllDemands.Value = 0
   FlxGridConfigure flxChildDemands
'
   If iIncDec > 0 Then
      szLastIDClicked = flxDemands.TextMatrix(flxDemands.Row, 1)
      szCurCompName = flxDemands.TextMatrix(flxDemands.Row, 2)
   Else
      szLastIDClicked = SelectedID()
   End If
   Call FillChildinGrid(szLastIDClicked, flxChildDemands)
   fraDetails.Caption = "Demand Details: " & szLastIDClicked
   iSelectedDemandsRow = iSelectedDemandsRow + iIncDec
End Sub

Private Sub flxEditDemand_Click()
   Dim iRow, iLeft, iCol As Integer
'
   iLeft = 0
   For iCol = 0 To flxEditDemand.ColSel - 1
      iLeft = iLeft + flxEditDemand.ColWidth(iCol)
   Next iCol
   Select Case flxEditDemand.ColSel
      Case 2:
         cboEditType.Left = iLeft + flxEditDemand.Left + 20
         cboEditType.Top = flxEditDemand.Top + (flxEditDemand.RowHeight(flxEditDemand.Row) * _
                                 flxEditDemand.Row) + 20
         cboEditType.Width = flxEditDemand.ColWidth(flxEditDemand.ColSel)
         cboEditType.Visible = True
         cboEditType.ZOrder 0
         cboEditType.SetFocus

      Case 4:                   'Description
         txtEditDescription.Left = iLeft + flxEditDemand.Left + 40
         txtEditDescription.Top = flxEditDemand.Top + (flxEditDemand.RowHeight(flxEditDemand.Row) * _
                                          flxEditDemand.Row) + 40
         txtEditDescription.Width = flxEditDemand.ColWidth(flxEditDemand.ColSel) - 40
         txtEditDescription.Height = flxEditDemand.RowHeightMin * 2
         txtEditDescription.Alignment = vbLeftJustify
         txtEditDescription.text = flxEditDemand.TextMatrix(flxEditDemand.Row, flxEditDemand.ColSel)
         txtEditDescription.Visible = True
         txtEditDescription.ZOrder 0
         txtEditDescription.SetFocus
'
      Case 5:           'From Date
         dtEditDate.Left = iLeft + flxEditDemand.Left + 40
         dtEditDate.Top = flxEditDemand.Top + (flxEditDemand.RowHeight(flxEditDemand.Row) * _
                              flxEditDemand.Row) + 40
         dtEditDate.Value = flxEditDemand.TextMatrix(flxEditDemand.Row, flxEditDemand.Col)
         dtEditDate.Visible = True
         dtEditDate.SetFocus
         dtEditDate.ZOrder 0
'
      Case 6:           'To Date
         dtEditDate.Left = iLeft + flxEditDemand.Left + 40
         dtEditDate.Top = flxEditDemand.Top + (flxEditDemand.RowHeight(flxEditDemand.Row) * _
                              flxEditDemand.Row) + 40
         dtEditDate.Visible = True
         dtEditDate.SetFocus
         dtEditDate.ZOrder 0
'
      Case 7:           'From Date
         dtEditDate.Left = iLeft + flxEditDemand.Left + 40
         dtEditDate.Top = flxEditDemand.Top + (flxEditDemand.RowHeight(flxEditDemand.Row) * _
                        flxEditDemand.Row) + 40
         dtEditDate.Visible = True
         dtEditDate.SetFocus
         dtEditDate.ZOrder 0
'
      Case 8:                          'Amount text box
         If flxEditDemand.TextMatrix(flxEditDemand.Rows - 1, 2) = "" Then
            flxEditDemand.Col = 2
            flxEditDemand_Click
            Exit Sub
         End If
         txtEditAmount.Left = iLeft + flxEditDemand.Left + 40
         txtEditAmount.Top = flxEditDemand.Top + (flxEditDemand.RowHeight(flxEditDemand.Row) * _
                                 flxEditDemand.Row) + 40
         txtEditAmount.Width = flxEditDemand.ColWidth(flxEditDemand.ColSel) - 40
         txtEditAmount.Alignment = vbRightJustify
         txtEditAmount.text = flxEditDemand.TextMatrix(flxEditDemand.Row, flxEditDemand.ColSel)
         txtEditAmount.Visible = True
         txtEditAmount.ZOrder 0
         txtEditAmount.SetFocus
'
      Case 18:                            'VAT Code
         cboEditVatCode.Left = iLeft + flxEditDemand.Left + 20 '- flxEditDemand.ColWidth(1)
         cboEditVatCode.Top = flxEditDemand.Top + (flxEditDemand.RowHeight(flxEditDemand.Row) * flxEditDemand.Row) + 20
         cboEditVatCode.Width = flxEditDemand.ColWidth(flxEditDemand.ColSel)
         cboEditVatCode.Visible = True
         cboEditVatCode.ZOrder 0
         cboEditVatCode.SetFocus
'
   End Select
End Sub

Private Sub flxPay_Click()
'   Select Case flxPay.Col
End Sub

Private Sub SPFlxGridConfigure()
   Dim szHeader As String
'
   flxSPayment.Clear
   flxSPayment.Cols = 12
   flxSPayment.Rows = 2
'
   szHeader$ = "<No.|<Type|<Tenant A/C|<Unit ID|<Date" & _
               "|<Ref|<Details|>Amount £|>O/S Amt. £" & _
               "|>Receipt £|>Discount|<DemandID"
   flxSPayment.FormatString = szHeader$
'
   flxSPayment.ColWidth(0) = 700    'No
   flxSPayment.ColWidth(1) = 800    'Type
   flxSPayment.ColWidth(2) = 1300   'Tenant A/c
   flxSPayment.ColWidth(3) = 1000   'Unit ID
   flxSPayment.ColWidth(4) = 1000   'Date
   flxSPayment.ColWidth(5) = 1800   'Ref
   flxSPayment.ColWidth(6) = 2200   'Details
   flxSPayment.ColWidth(7) = 1000   'Amount
   flxSPayment.ColWidth(8) = 1000   'O/S Amount
   flxSPayment.ColWidth(9) = 1000   'Receipt
   flxSPayment.ColWidth(10) = 0     'Discount
   flxSPayment.ColWidth(11) = 0     'DemandID
'
   flxSPayment.RowHeightMin = 285
End Sub

Private Sub flxListBk_DblClick(Index As Integer)
   If sTextBox = "Bank" Then
      txtBkAc(tabPayment.Tab - 1).text = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 0)
      cmdTaxListBk(tabPayment.Tab - 1).Caption = "T1"
      nTaxCode = TaxRate(1)
'
      txtBkAc(tabPayment.Tab - 1).SetFocus
   End If
   If sTextBox = "Unit" Then
      txtUnitBk(tabPayment.Tab - 1).text = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 0)
   End If
   If sTextBox = "NC" Then
      txtNCBk(tabPayment.Tab - 1).text = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 0)
   End If
   If sTextBox = "Dept" Then
      txtDeptBk(tabPayment.Tab - 1).text = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 0)
   End If
   If sTextBox = "Proj" Then
      txtProjBk(tabPayment.Tab - 1).text = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 0)
   End If
   If sTextBox = "CC" Then
      txtCCBk(tabPayment.Tab - 1).text = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 0)
   End If
   If sTextBox = "VAT" Then
      nTaxCode = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 1)
      cmdTaxListBk(tabPayment.Tab - 1).Caption = flxListBk(tabPayment.Tab - 1).TextMatrix(flxListBk(tabPayment.Tab - 1).Row, 0)
      txtVatBk(tabPayment.Tab - 1).text = Format(IIf(txtNetBk(tabPayment.Tab - 1).text = "", 0, Val(txtNetBk(tabPayment.Tab - 1).text)) * _
                     (nTaxCode / 100), "0.00")
   End If

   flxListBk(tabPayment.Tab - 1).Clear
   flxListBk(tabPayment.Tab - 1).Cols = 2
   flxListBk(tabPayment.Tab - 1).Rows = 2
   fraListBk(tabPayment.Tab - 1).Visible = False
End Sub

Private Sub flxListNC_DblClick()
   txtNominalCodeTR.text = flxListNC.TextMatrix(flxListNC.Row, 0)
   flxListNC.Clear
   flxListNC.Cols = 2
   flxListNC.Rows = 2
   fraListNC.Visible = False
End Sub

Private Sub flxListNC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then imgFramListCoseNC_Click
End Sub

Private Sub flxSPayment_Click()
   Dim i As Integer
   Dim iSPFlxLeft As Integer
'
   iSPFlxLeft = flxSPayment.Left
   For i = 0 To flxSPayment.Col - 1
      iSPFlxLeft = iSPFlxLeft + flxSPayment.ColWidth(i)
   Next i
'
   Select Case flxSPayment.Col
   Case Is = 9
      If flxSPayment.TextMatrix(flxSPayment.Row, 2) = "" Then Exit Sub
      txtSPayment.Top = flxSPayment.Top + (flxSPayment.RowHeight(flxSPayment.Row) * flxSPayment.Row) + 15
      txtSPayment.Left = iSPFlxLeft + 10
      txtSPayment.Width = flxSPayment.ColWidth(flxSPayment.Col)
      txtSPayment.text = flxSPayment.TextMatrix(flxSPayment.Row, flxSPayment.Col)
      txtSPayment.ZOrder 0
      txtSPayment.Visible = True
      txtSPayment.SetFocus

'   Case Is = 10
'      If flxSPayment.TextMatrix(flxSPayment.Row, 2) = "" Then Exit Sub
'      txtSPayment.Top = flxSPayment.Top + (flxSPayment.RowHeight(flxSPayment.Row) * flxSPayment.Row) + 15
'      txtSPayment.Left = iSPFlxLeft + 10
'      txtSPayment.Width = flxSPayment.ColWidth(flxSPayment.Col)
'      txtSPayment.text = flxSPayment.TextMatrix(flxSPayment.Row, flxSPayment.Col)
'      txtSPayment.ZOrder 0
'      txtSPayment.Visible = True
'      txtSPayment.SetFocus
   End Select
End Sub

Private Sub Form_Load()
   MousePointer = vbHourglass
'
   iIncDec = 0
   Me.Top = 50
   Me.Left = 50
   bChangesMade = False
   nTaxCode = 17.5
   iSelected = 0
   bChangesMade = False
'
   Dim a As Integer
'
   Me.Caption = "Demands"
'
   tabDmdRcpt.Tab = 0
'
   FrameConfigure
'
   FlxDemandsConfigure flxDemandHistory
   FillDemandsFlxGrid flxDemandHistory, True    'True - uploading history records
'
   FlxDemandsConfigure flxDemands
   FillDemandsFlxGrid flxDemands, False         'Flase - uploading history, which are already posted and printed and exported to sage
'
   FlxGridConfigure flxChildDemands
   FlxGridConfigure flxChildDemandHistory
'
   Call LoadFreq      'Load all frequensis in the array string
'
'Define Bank Payment flex grid
   BankFlxGridConfigure flxBankPay(0)
   BankFlxGridConfigure flxBankPay(1)
'
   BankFlxGridLoad flxBankPay(0)
   FlxSumUp flxBankPay(0), 10, 12, txtBkTotalNet(0), txtBkTotalVat(0)
   BankFlxGridLoad flxBankPay(1)
   FlxSumUp flxBankPay(1), 10, 12, txtBkTotalNet(1), txtBkTotalVat(1)
'
   fVAT_Rate = VatRate
'
   SPFlxGridConfigure
'
'Bring all Invoices or Demands into tlbReceipt table
   MigrateInvintoReceipt
'
   flxDemands.Row = 0
   flxDemands.Col = 0
   
   MousePointer = vbDefault
End Sub

Private Sub BankFlxGridLoad(conFlxGrid As Control)
   Dim iRow As Integer
   Dim szStr As String
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
   Dim szBPBR As String
'
   szBPBR = IIf(conFlxGrid.Index = 0, "BP", "BR")
'
   Set adoConn = New ADODB.Connection
   Set adoRst = New ADODB.Recordset
'
'   connect to database
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'
   szStr = "SELECT tlbBankPayment.MY_ID, tlbBankPayment.TRAN_ID, " & _
                  "tlbBankPayment.BANK_AC, tlbBankPayment.TRAN_DATE, " & _
                  "tlbBankPayment.TRANS, tlbBankPayment.TRAN_TYPE, " & _
                  "tlbBankPayment.UNIT_ID, " & _
                  "tlbBankPayment.NOMINAL_CODE, tlbBankPayment.DEPT_ID, " & _
                  "tlbBankPayment.PROJ_REF, tlbBankPayment.COST_CODE, " & _
                  "tlbBankPayment.DESCRIPTION, tlbBankPayment.NET_AMOUNT, " & _
                  "tlbBankPayment.TAX_CODE, tlbBankPayment.VAT, " & _
                  "tlbBankPayment.UPDATE_SAGE, tlbBankPayment.RECHARABLE " & _
            "FROM tlbBankPayment " & _
            "WHERE tlbBankPayment.UPDATE_SAGE = FALSE AND " & _
                  "tlbBankPayment.TRAN_TYPE = '" & szBPBR & "' " & _
            "ORDER BY tlbBankPayment.TRAN_ID;"
   adoRst.Open szStr, adoConn, adOpenDynamic, adLockPessimistic
   
   iRow = 1
   While Not adoRst.EOF
      conFlxGrid.TextMatrix(iRow, 0) = adoRst!BANK_AC
      conFlxGrid.TextMatrix(iRow, 1) = adoRst!TRAN_DATE
      conFlxGrid.TextMatrix(iRow, 2) = IIf(IsNull(adoRst!TRANS), "", adoRst!TRANS)
      conFlxGrid.TextMatrix(iRow, 3) = IIf(IsNull(adoRst!TRAN_TYPE), "", adoRst!TRAN_TYPE)
      conFlxGrid.TextMatrix(iRow, 4) = IIf(IsNull(adoRst!UNIT_ID), "", adoRst!UNIT_ID)
      conFlxGrid.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!NOMINAL_CODE), "", adoRst!NOMINAL_CODE)
      conFlxGrid.TextMatrix(iRow, 6) = IIf(IsNull(adoRst!DEPT_ID), "", adoRst!DEPT_ID)
      conFlxGrid.TextMatrix(iRow, 7) = IIf(IsNull(adoRst!PROJ_REF), "", adoRst!PROJ_REF)
      conFlxGrid.TextMatrix(iRow, 8) = IIf(IsNull(adoRst!COST_CODE), "", adoRst!COST_CODE)
      conFlxGrid.TextMatrix(iRow, 9) = IIf(IsNull(adoRst!description), "", adoRst!description)
      conFlxGrid.TextMatrix(iRow, 10) = IIf(IsNull(adoRst!NET_AMOUNT), "", adoRst!NET_AMOUNT)
      conFlxGrid.TextMatrix(iRow, 11) = IIf(IsNull(adoRst!TAX_CODE), "", adoRst!TAX_CODE)
      conFlxGrid.TextMatrix(iRow, 12) = IIf(IsNull(adoRst!VAT), "", adoRst!VAT)
      conFlxGrid.TextMatrix(iRow, 13) = IIf(adoRst!RECHARABLE, "X", "")
      conFlxGrid.TextMatrix(iRow, 14) = adoRst!MY_ID
      conFlxGrid.TextMatrix(iRow, 16) = "0"
      adoRst.MoveNext
      If Not adoRst.EOF Then conFlxGrid.AddItem ""
      iRow = iRow + 1
   Wend
'
   adoRst.Close
   adoConn.Close
'
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub FlxGridConfigure(conFlxGrid As Control)
   conFlxGrid.Cols = 18
'
   If conFlxGrid.Rows = 2 Then
      conFlxGrid.ColWidth(0) = 250     'Solid column
      conFlxGrid.ColWidth(1) = 550     'Split ID
      conFlxGrid.ColWidth(2) = 500     'Generate Demand (A/M)
      conFlxGrid.ColWidth(3) = 4200    'Description
      conFlxGrid.ColWidth(4) = 1000    'From Date
      conFlxGrid.ColWidth(5) = 1000    'To Date
      conFlxGrid.ColWidth(6) = 1000    'Due Date
      conFlxGrid.ColWidth(7) = 1000    'Amount
      conFlxGrid.ColWidth(8) = 1000    'VAT
      conFlxGrid.ColWidth(9) = 1000    'Total
      conFlxGrid.ColWidth(10) = 0       'NC Amt
      conFlxGrid.ColWidth(11) = 0       'NN Amt
      conFlxGrid.ColWidth(12) = 0      'NC VAT
      conFlxGrid.ColWidth(13) = 0      'NN VAT
      conFlxGrid.ColWidth(14) = 0      'NC Tol
      conFlxGrid.ColWidth(15) = 0      'NN Tol
      conFlxGrid.ColWidth(16) = 0      'BATCHID number
      conFlxGrid.ColWidth(17) = 1000    'SageRef
   End If
   conFlxGrid.Rows = 2
   conFlxGrid.Clear
'
   conFlxGrid.TextMatrix(0, 1) = "S_ID"
   conFlxGrid.TextMatrix(0, 2) = "A/M"
   conFlxGrid.TextMatrix(0, 3) = "Description"
   conFlxGrid.TextMatrix(0, 4) = "From Dt"
   conFlxGrid.TextMatrix(0, 5) = "To Dt"
   conFlxGrid.TextMatrix(0, 6) = "Due Dt"
   conFlxGrid.TextMatrix(0, 7) = "Amount"
   conFlxGrid.TextMatrix(0, 8) = "VAT"
   conFlxGrid.TextMatrix(0, 9) = "Total"
   conFlxGrid.TextMatrix(0, 17) = "SageRef"
'
   conFlxGrid.RowHeightMin = 315
End Sub

Private Sub FlxGridManualConfigure(conFlxGrid As Control)
   conFlxGrid.Cols = 19
'
   If conFlxGrid.Rows = 2 Then
      conFlxGrid.ColWidth(0) = 250     'Solid column
      conFlxGrid.ColWidth(1) = 490     'Split ID
      conFlxGrid.ColWidth(2) = 1480    'Demand Type
      conFlxGrid.ColWidth(3) = 380     'Generate Demand (A/M)
      conFlxGrid.ColWidth(4) = 2880    'Description
      conFlxGrid.ColWidth(5) = 840     'From Date
      conFlxGrid.ColWidth(6) = 840     'To Date
      conFlxGrid.ColWidth(7) = 840     'Due Date
      conFlxGrid.ColWidth(8) = 840     'Amount
      conFlxGrid.ColWidth(9) = 740     'VAT
      conFlxGrid.ColWidth(10) = 840    'Total
      conFlxGrid.ColWidth(11) = 0      'NC Amt
      conFlxGrid.ColWidth(12) = 0      'NN Amt
      conFlxGrid.ColWidth(13) = 0      'NC VAT
      conFlxGrid.ColWidth(14) = 0      'NN VAT
      conFlxGrid.ColWidth(15) = 0      'NC Tol
      conFlxGrid.ColWidth(16) = 0      'NN Tol
      conFlxGrid.ColWidth(17) = 880    'SageRef
      conFlxGrid.ColWidth(18) = 680    'Vat Code
   End If
   conFlxGrid.Rows = 2
   conFlxGrid.Clear
'
   conFlxGrid.TextMatrix(0, 1) = "SN"
   conFlxGrid.TextMatrix(0, 2) = "Demand Type"
   conFlxGrid.TextMatrix(0, 3) = "A/M"
   conFlxGrid.TextMatrix(0, 4) = "Description"
   conFlxGrid.TextMatrix(0, 5) = "From Dt"
   conFlxGrid.TextMatrix(0, 6) = "To Dt"
   conFlxGrid.TextMatrix(0, 7) = "Due Dt"
   conFlxGrid.TextMatrix(0, 8) = "Amount"
   conFlxGrid.TextMatrix(0, 9) = "VAT"
   conFlxGrid.TextMatrix(0, 10) = "Total"
   conFlxGrid.TextMatrix(0, 17) = "SageRef"
   conFlxGrid.TextMatrix(0, 18) = "V/C"
'
   conFlxGrid.RowHeightMin = 315
End Sub

Private Sub LoadFreq()
   Dim adoRstFreq As ADODB.Recordset
   Dim adoConn As ADODB.Connection
   Dim strSQLTitles As String
'
   Set adoConn = New ADODB.Connection
   Set adoRstFreq = New ADODB.Recordset
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   strSQLTitles = "SELECT * FROM FREQUENCIES;"
   adoRstFreq.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
'
   ReDim szaFreq(adoRstFreq.RecordCount) As String
'
   While Not adoRstFreq.EOF
      szaFreq(adoRstFreq.Fields("ID").Value) = adoRstFreq.Fields("CALDAYS").Value
      adoRstFreq.MoveNext
   Wend
   adoRstFreq.Close
   adoConn.Close
   Set adoRstFreq = Nothing
   Set adoConn = Nothing
End Sub

Private Sub FrameConfigure()
   fraGenerate.Height = fraMain.Height
End Sub

Private Sub FillDemandsFlxGrid(conFlxGrid As Control, bHistory As Boolean)
   Dim iRow As Integer
   Dim szStr As String
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
'
   Set adoConn = New ADODB.Connection
   Set adoRst = New ADODB.Recordset
'
'   connect to database
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'***get sendtoprint field for all demands that were set to be sent to print***
   szStr = "SELECT SUM(DEMANDSPLITRECORDS.TOTALAMOUNT) AS TOTAL, " & _
                  "DEMANDRECORDS.SageAccountNumber AS SAGE_AC, " & _
                  "DEMANDRECORDS.TenantCompanyName AS T_NAME, " & _
                  "DEMANDRECORDS.UnitNumber AS U_NUM, " & _
                  "DEMANDRECORDS.TransactionType as T_TP, " & _
                  "DEMANDRECORDS.ISSUEDATE AS I_DATE, " & _
                  "DEMANDRECORDS.IsPrinted AS I_PRINT, " & _
                  "DEMANDRECORDS.UPDATE_SAGE AS E_SAGE, " & _
                  "DEMANDRECORDS.DemandID AS D_ID, " & _
                  "DEMANDRECORDS.DemandHistory AS D_HIST, " & _
                  "DEMANDRECORDS.BATCHID AS BT " & _
           "FROM DEMANDRECORDS, DEMANDSPLITRECORDS " & _
           "WHERE DEMANDRECORDS.DEMANDID=DEMANDSPLITRECORDS.DEMANDID AND " & _
                  "DEMANDRECORDS.DEMANDHISTORY = " & CBool(bHistory) & " " & _
           "GROUP BY DEMANDRECORDS.SageAccountNumber, " & _
                  "DEMANDRECORDS.TenantCompanyName, " & _
                  "DEMANDRECORDS.UnitNumber, " & _
                  "DEMANDRECORDS.TransactionType, " & _
                  "DEMANDRECORDS.ISSUEDATE, " & _
                  "DEMANDRECORDS.IsPrinted, " & _
                  "DEMANDRECORDS.UPDATE_SAGE, " & _
                  "DEMANDRECORDS.DemandID, " & _
                  "DEMANDRECORDS.DemandHistory, " & _
                  "DEMANDRECORDS.BATCHID " & _
           "ORDER BY DEMANDRECORDS.DemandID"
'
'Debug.Print szStr
   adoRst.Open szStr, adoConn, adOpenDynamic, adLockPessimistic
'
   If Not adoRst.EOF Then
      iRow = 1
      While Not adoRst.EOF
'         If iRow Mod 2 = 0 Then
'             conFlxGrid.CellBackColor = &HF4F4F4
'         Else
'             conFlxGrid.CellBackColor = &H80000014
'         End If
'
         conFlxGrid.TextMatrix(iRow, 1) = Format(adoRst!D_ID, "00000000")
         conFlxGrid.TextMatrix(iRow, 2) = IIf(adoRst!T_TP = 1, "INV", "CRN")
         conFlxGrid.TextMatrix(iRow, 3) = adoRst!T_NAME
         conFlxGrid.TextMatrix(iRow, 4) = adoRst!U_NUM
         conFlxGrid.TextMatrix(iRow, 5) = Format(adoRst!I_DATE, "dd/mm/yyyy")
         conFlxGrid.TextMatrix(iRow, 6) = Format(adoRst.Fields("TOTAL").Value, "0.00")
         conFlxGrid.TextMatrix(iRow, 7) = IIf(adoRst!I_PRINT, "YES", "NO")
         conFlxGrid.TextMatrix(iRow, 8) = adoRst!SAGE_AC
         conFlxGrid.TextMatrix(iRow, 9) = IIf(adoRst!E_SAGE, "YES", "NO")
         conFlxGrid.TextMatrix(iRow, 10) = adoRst!BT
         adoRst.MoveNext
         If Not adoRst.EOF Then conFlxGrid.AddItem ""
         iRow = iRow + 1
      Wend
   Else
      If Not bHistory Then MsgBox "There are no demands!", vbOKOnly + vbInformation, "Demands"
      If bHistory Then MsgBox "There no demand history!", vbOKOnly + vbInformation, "Demands"
   End If
'
   adoRst.Close
   adoConn.Close
'
   Set adoRst = Nothing
   Set adoConn = Nothing
   
   iSelectedDemandsRow = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
End Sub

Private Sub mnuDeleteDemand_Click()
'   Call Delete
End Sub

Private Sub mnuEditDemands_Click()

Call Edit

End Sub

Private Sub mnuExit_Click()
   Unload frmMMain
End Sub

Private Sub mnuGenAll_Click()

'    Call GenAutoDemands

End Sub

Private Sub mnuGenManual_Click()

    Call AddManualDemand

End Sub

Public Sub AddManualDemand()

'   Call FillcboTenant
'
''   Label1.Visible = True
'   cboTenant.Visible = True
'   Call EmptyBoxes
'   chk1.Visible = False
'   chk2.Visible = False
'   chk3.Visible = False
'
'   cmdEdit.Visible = False
'   cmdGenerateManual.Visible = False
'   cmdDelete.Visible = False
'   cmdDeleteOld.Visible = False
'   cmdCancelEdit.Visible = False
'   cmdSaveNew.Visible = False
'   cmdCancelNew.Visible = False
'   cmdGenAll.Visible = False
'   cmdPrint.Visible = False
'   cmdPrintThis.Visible = False
'   cmdPrintBatch.Visible = False
'   cmdReprint.Visible = False
''    mnuEditDemands.Enabled = False
''    mnuGenManual.Enabled = False
''    mnuGenAll.Enabled = False
''    mnuPrint.Enabled = False
''    mnuReprint.Enabled = False
'   cmdMoveFirst.Visible = False
'   cmdMoveLast.Visible = False
'   cmdMovePrevious.Visible = False
'   cmdMoveNext.Visible = False
''   lbl1.Visible = False
''   lbl2.Visible = False
''   lbl3.Visible = False
''   lbl4.Visible = False
''   cmdFind.Visible = False
''   Text2.Visible = False

End Sub

Public Sub EnableBoxes()

'    txt4.Enabled = True
'    txtDueDate.Enabled = True
'    txtRef.Enabled = True
'    txtAmount.Enabled = True
'    txtDescription.Enabled = True
'    txtSageText.Enabled = True
'    cboType.Enabled = True

End Sub

Public Sub DisableBoxes()
    
'    txt4.Enabled = False
'    txtDueDate.Enabled = False
'    txtRef.Enabled = False
'    txtAmount.Enabled = False
'    txtDescription.Enabled = False
'    txtSageText.Enabled = False
'    cboType.Enabled = False

End Sub

Public Sub EmptyBoxes()

'Text1.text = ""
'txtCompanyName.text = ""
'txtSageAccNumber.text = ""
'txt3.text = ""
'txt4.text = ""
'txtDueDate.text = ""
'txtRef.text = ""
'txtAmount.text = ""
'txtVatAmt.text = ""
'txtTotal.text = ""
'txtDescription.text = ""
'txtSageText.text = ""
'cboDemand.text = ""
'cboType.text = ""

End Sub

Public Sub FillcboTenant()
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String
   
   cboTenant.Clear

   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt

   SQLStr1 = "SELECT Tenants.CompanyName as CN, Tenants.SageAccountNumber as SAN, Units.UnitNumber " & _
             "FROM Tenants, LeaseDetails, Units " & _
             "WHERE Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                   "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                   "Units.Occupied = 'Y' " & _
             "ORDER BY Tenants.SageAccountNumber"
'Debug.Print SQLStr1
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

   If rdoRst1.EOF = False Then
       While rdoRst1.EOF = False
           'If rdoRst1!CurrentRental <> "" Then cboTenant.AddItem rdoRst1!SageAccountNumber & " / " & rdoRst1!CompanyName
           cboTenant.AddItem rdoRst1!SAN & " / " & rdoRst1!CN
           rdoRst1.MoveNext
       Wend
   End If

   rdoRst1.Close
   rdoConn.Close
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Sub

'*********************
'   FillCbos() collects all type of Demond and push into cboType Combo Box.
'*********************
Public Sub FillCbos()
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String
'
'    cboDemand.Enabled = True
'    cboDemand.AddItem "Manual", 0
'    cboDemand.AddItem "Automatic", 1
'    cboDemand.Enabled = False
'
    rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
    rdoConn.CursorDriver = rdUseIfNeeded
    rdoConn.EstablishConnection rdDriverNoPrompt
'
    SQLStr1 = "SELECT ID, Type FROM DemandTypes"
    Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
    If rdoRst1.EOF = False Then
        While rdoRst1.EOF = False
            cboType.AddItem rdoRst1!ID & " / " & rdoRst1!Type
            rdoRst1.MoveNext
        Wend
    End If
'
    rdoRst1.Close
    rdoConn.Close
'
End Sub

Private Sub mnuGlobal_Click()

'Load frmGlobal
'frmGlobal.Show
'Unload Me

End Sub

Private Sub mnuLease_Click()
'
'Load frmLease
'frmLease.Show
'Unload Me

End Sub

Private Sub mnuMain_Click()

    'Load frmMain
    'frmMain.Show
    frmMMain.fraCmdButton.Enabled = True
    Unload Me

End Sub

Private Sub mnuPrint_Click()

Call PrintDemands

End Sub

Private Sub mnuPrintBatch_Click()

Call PrintBatchSelected

End Sub

Private Sub mnuShopCentre_Click()

'Load frmShoppingCentre
'frmShoppingCentre.Show
'Unload Me

End Sub

Private Sub mnuTenants_Click()

'    Load frmTenant
'    frmTenant.Show
'    Unload Me

End Sub

Private Sub mnuUnits_Click()
'    Load frmUnit
'    frmUnit.Show
'    Unload Me
End Sub

Private Sub imgFramListCoseBk_Click(Index As Integer)
   fraListBk(tabPayment.Tab - 1).Visible = False
End Sub

Private Sub imgFramListCoseNC_Click()
   fraListNC.Visible = False
End Sub

'Private Sub Image1_Click()
'   fraPayee.Visible = False
''   tabDmdRcpt.Enabled = True
'End Sub

Private Sub lblDeleteDemands_Click()
   If fraDeleteDemand.Top = 1920 Then Exit Sub

   fraDeleteDemand.Top = 1560
   fraEditDemand.Top = 1080
   fraReprintDemands.Top = 600
'   fraPrintDemands.Top = 480
End Sub

Private Sub lblDeleteDemands_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblDeleteDemands.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblDeleteDemands_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblDeleteDemands.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblEditDemand_Click()
'   If fraEditDemand.Top = 1080 Then
'      fraDeleteDemand.Top = 5400
'      Exit Sub
'   End If

'   fraDeleteDemand.Top = 1920
   fraEditDemand.Top = 480
'   fraReprintDemands.Top = 600
'   fraPrintDemands.Top = 480
'   fraEditDemand.Top = 4920
'   fraReprintDemands.Top = 4440
'   fraPrintDemands.Top = 3960
End Sub

Private Sub lblEditDemand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblEditDemand.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblEditDemand_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblEditDemand.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblGenerate_Click()
   If cmdEdit.Enabled = False Then Exit Sub
'   fraDeleteDemand.Top = 5400
   fraEditDemand.Top = 2400
'   fraReprintDemands.Top = 4440
'   fraPrintDemands.Top = 3960
End Sub

Private Sub lblGenerate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblGenerate.MouseIcon = LoadPicture(App.Path + "\Package1\hmove.cur")
End Sub

Private Sub lblGenerate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblGenerate.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblPrintDemands_Click()
'   If fraPrintDemands.Top = 480 Then
'      fraDeleteDemand.Top = 5400
'      fraEditDemand.Top = 4920
'      fraReprintDemands.Top = 4440
'      Exit Sub
'   End If
'   fraPrintDemands.Top = 480
End Sub

Private Sub lblPrintDemands_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   lblPrintDemands.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub lblPrintDemands_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   lblPrintDemands.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub lblReprintDemand_Click()
   If fraReprintDemands.Top = 600 Then
      fraDeleteDemand.Top = 5400
      fraEditDemand.Top = 4920
      Exit Sub
   End If
   fraReprintDemands.Top = 600
'   fraPrintDemands.Top = 480
End Sub

Private Sub lblReprintDemand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblReprintDemand.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
End Sub

Private Sub txt4_LostFocus()
'    If CheckDate(txt4.text) = False Then txt4.text = ""
    Call GetReference
End Sub

Private Sub lstTypeBk_DblClick(Index As Integer)
   txtTypeBk(tabPayment.Tab - 1).text = lstTypeBk(tabPayment.Tab - 1).text
   lstTypeBk(tabPayment.Tab - 1).Visible = False
   txtTypeBk(tabPayment.Tab - 1).SetFocus
End Sub

Private Sub lstTypeBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then lstTypeBk_DblClick (tabPayment.Tab - 1)
End Sub

Private Sub lstYNBk_Click(Index As Integer)
   txtRecharge(tabPayment.Tab - 1).text = lstYNBk(tabPayment.Tab - 1).text
   lstYNBk(tabPayment.Tab - 1).Visible = False
   txtRecharge(tabPayment.Tab - 1).SetFocus
End Sub

Private Sub tabDmdRcpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub tabPayment_Click(PreviousTab As Integer)
   If tabPayment.Tab = 0 Or tabPayment.Tab = 3 Then Exit Sub
   If cmdNewBk(tabPayment.Tab - 1).Enabled And cmdEditBk(tabPayment.Tab - 1).Enabled Then
      HandleTextBoxesBk True, False
   End If
End Sub

Private Sub tabPayment_LostFocus()
   Select Case tabPayment.Tab
      Case 0: cmbBankAc.SetFocus
      Case 1: cmdNewBk(tabPayment.Tab - 1).SetFocus
      Case 2: cmdNewBk(tabPayment.Tab - 1).SetFocus
   End Select
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub txtAddNewDescription_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then txtAddNewDescription_LostFocus
End Sub

Private Sub txtDueDate_Click()
'   dtDueDate.Visible = True
'   dtDueDate.ZOrder 0
End Sub

Private Sub txtAmount_Change()
'    If txtAmount.text <> "" And txtAmount.text <> "-" Then
'        If NumberCheck2(txtAmount.text) = False Then
'            txtAmount.text = ""
'        Else
'            If DemandId = 4 Then
'                txtTotal.text = txtAmount.text
'                txtVatAmt.text = "0"
'            Else
'                txtVatAmt.text = Round(CDbl(txtAmount.text) * VatRate / 100, 2)
'                txtTotal.text = CDbl(txtAmount.text) + CDbl(txtVatAmt.text)
'            End If
'        End If
'    End If
End Sub

Public Sub GetRecord()
'   Dim adoConn As adodb.Connection
'   Dim adoRst As adodb.Recordset
'   Dim SQLStr1 As String
''   txtCompanyName.text = adoRst!TenantCompanyName
''   txtSageAccNumber.text = adoRst!SageAccountNumber
''   txt3.text = adoRst!UnitNumber
''   Text1.text = adoRst!UniqueRefNumber
''   If adoRst!AutomaticManual = "A" Then cboDemand.text = "Automatic"
''   If adoRst!AutomaticManual = "M" Then cboDemand.text = "Manual"
''   txt4.text = adoRst!IssueDate
''   txtDueDate.text = adoRst!DueDate
'   If EditMode Then typeofdemand = adoRst!typeofdemand
'   cboType.text = cboType.List(adoRst!typeofdemand)
'
'   'that is a very messy way of doing this...code added by KC 11:01 9th October
''   Call cboType_LostFocus
'
''   txtAmount.text = adoRst!Amount
''   txtVatAmt.text = adoRst!VATAmount
''   txtTotal.text = adoRst!TotalAmount
'   txtDescription.text = adoRst!description
'   txtSageText.text = adoRst!text
'   If adoRst!IsPrinted = "Y" Then chk1.Value = 1 And chkPrint.Value = 0
'   If adoRst!IsPrinted = "N" Then chk1.Value = 0 And chkPrint.Value = 0
'   If adoRst!SendToPrint = "Y" Then chkPrint.Value = 1 Else chkPrint.Value = 0
'   If adoRst!UPDATE_SAGE = "Y" Then chk2.Value = 1
'   If adoRst!UPDATE_SAGE = "N" Then chk2.Value = 0
'   If adoRst!ExportedToExcel = "Y" Then chk3.Value = 1
'   If adoRst!ExportedToExcel = "N" Then chk3.Value = 0
End Sub

Public Sub GetReference()
'
'If txt4.text <> "" Then
'    If cboType.text <> "" Then
'        txtRef.text = Prefix & Left(txt4.text, 2) & Mid(txt4.text, 4, 2) & Right(txt4.text, 2)
'    End If
'End If

End Sub

Public Sub Edit()
   Dim szComp As String, szUnit As String, szBt As String
   Dim szSageAC As String, dtIssueDt As Date
'
   FlxGridManualConfigure flxEditDemand
'
   If cboEditType.ListCount = 0 Then Call FillcboType(cboEditType, InvCre(szLastIDClicked))
   If cboEditVatCode.ListCount = 0 Then Call FillcboVatCode(cboEditVatCode)
'
   txtEditDemandID.text = Format(flxDemands.TextMatrix(flxDemands.RowSel, 1), "00000000")
'
   Call FindLastSelID(szComp, szUnit, szSageAC, dtIssueDt, szBt)
   szCurCompSageAccNum = szSageAC
   txtEditBatch.text = szBt
'
   txtEditTenantName.text = szSageAC & " / " & szComp
   txtEditIssueDate.text = Format(dtIssueDt, "dd mmmm yyyy")
   txtEditAddNewSageText.text = "S/L " & szSageAC
'
   txtEditUnit.text = szUnit
'*****Fill grid by selected demands****
   Call FillManualChildinGrid(szLastIDClicked, szBt, flxEditDemand)
   lblTransactionType.Caption = "Transaction type: " & flxDemands.TextMatrix(flxDemands.Row, 2)
'
'Calculate sub total to bottom of the grid
   CalSubTotal flxEditDemand, txtEditSubAmount, txtEditSubVat, txtEditSubTotal
End Sub

Private Function PartString(szString As String, iStringPartNum As Integer, szDelemeter As String) As String
   Dim szaString() As String
   
   szaString = Split(szString, szDelemeter)
   PartString = szaString(iStringPartNum)
End Function

Private Sub GenAutoConDemands()
   Dim szaNCode() As String, szaNCName() As String, szaPrefix() As String
   Dim szDes As String, CutOffDate As String, szSQLStr As String
'
   Dim BRcount As Integer, SCcount As Integer, IPcount As Integer, ICcount As Integer
   Dim NextUniqueRefNo As Long, a As Integer, iSerial As Integer
   Dim lBatch As Long, lDemand As Long, iChildId As Integer

   Dim dtEndDate As Date, dtNtDueDate As Date

   Dim adoRstDemandRec As ADODB.Recordset, adoRstDmdTyp As ADODB.Recordset
   Dim adoRstLeaseDtl As ADODB.Recordset, adoRstSplitDemand As ADODB.Recordset
   Dim adoRstRentCharges As ADODB.Recordset
   Dim adoConn As ADODB.Connection

   Dim DaysB4Due As Integer

   If MsgBox("  Are you sure you wish to generate automatic demands?" & (Chr(13) + Chr(10)) & _
             "Please ensure your global data has been correctly inputted.", vbYesNo + vbQuestion, _
             "Generate Automatic Demands") = vbNo Then Exit Sub

   On Error GoTo ErrH

   MousePointer = vbHourglass    'change the mouse cursor to show program is working
'
   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'
'   Connect to Demands table to add new demands.
   Set adoRstDemandRec = New ADODB.Recordset
   szSQLStr = "SELECT * FROM DemandRecords"
   adoRstDemandRec.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
'
   Set adoRstSplitDemand = New ADODB.Recordset
   szSQLStr = "SELECT * FROM DemandSplitRecords"
   adoRstSplitDemand.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
'
'   get nominal codes and prefix for base rent from demand types.
   Set adoRstDmdTyp = New ADODB.Recordset
   szSQLStr = "SELECT * FROM DemandTypes;"
   adoRstDmdTyp.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'
   If adoRstDmdTyp.EOF Then
      MsgBox "There are no Demand Type in the database", vbCritical, "DATA Input Error"
'
      adoRstSplitDemand.Close
      adoRstDemandRec.Close
      adoRstDmdTyp.Close
      adoConn.Close
      Set adoRstSplitDemand = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstDmdTyp = Nothing
      Set adoConn = Nothing
      Exit Sub
   End If

   ReDim szaPrefix(adoRstDmdTyp.RecordCount) As String
   ReDim szaNCName(adoRstDmdTyp.RecordCount) As String
   ReDim szaNCode(adoRstDmdTyp.RecordCount) As String

'**Saving all Nominal Codes and Prefixes in the array
   While Not adoRstDmdTyp.EOF
      szaPrefix(adoRstDmdTyp!ID) = adoRstDmdTyp!Prefix
      szaNCName(adoRstDmdTyp!ID) = adoRstDmdTyp!NominalNameforAmount & " # " & adoRstDmdTyp!NominalNameforVAT & " # " & adoRstDmdTyp!NominalNameforTotal
      szaNCode(adoRstDmdTyp!ID) = adoRstDmdTyp!NominalCodeforAmount & " # " & adoRstDmdTyp!NominalCodeForVAT & " # " & adoRstDmdTyp!NominalCodeForTotal
      adoRstDmdTyp.MoveNext
   Wend
   adoRstDmdTyp.Close
   Set adoRstDmdTyp = Nothing

   Set adoRstLeaseDtl = New ADODB.Recordset

   szSQLStr = "SELECT * FROM LeaseDetails WHERE Status = TRUE;"
   adoRstLeaseDtl.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   iSerial = 1
   lBatch = NextRef(adoConn, "BATCH_REF")
   If adoRstLeaseDtl.EOF Then
       adoRstLeaseDtl.Close
       Set adoRstLeaseDtl = Nothing
   Else
      While Not adoRstLeaseDtl.EOF

'*********************** SAMRAT 25/11/2005***************************************
'*** Determin the date boundray in future, in this boundary
'*** we have to find those demands' due date to calcuate
'*** demands from lease table and we have to calculate &
'*** collect all those demands in DemandRecords table.
'*********************************************************************************
         DaysB4Due = GlbDaysBeforeDue(adoRstLeaseDtl.Fields("LeaseID").Value)
         CutOffDate = DateAdd("d", DaysB4Due, Date) 'the Date function returns the current system date

         If (adoRstLeaseDtl.Fields("InterestChargeable").Value = "Y" And _
               Val(IIf(IsNull(adoRstLeaseDtl!InterestAmount), 0, adoRstLeaseDtl!InterestAmount)) > 0) Or _
            (adoRstLeaseDtl!BRPayable = "Y" And _
               GeneratableBaseRent(adoConn, adoRstLeaseDtl!LeaseID, DaysB4Due)) Or _
            (adoRstLeaseDtl!SCPayable = "Y" And DateDiff("d", _
               IIf(adoRstLeaseDtl!SCNextDueDate = "", _
               DateAdd("D", 1, CutOffDate), adoRstLeaseDtl!SCNextDueDate), _
               CutOffDate) >= 0 And Val(IIf(IsNull(adoRstLeaseDtl!SCAmount), 0, adoRstLeaseDtl!SCAmount)) > 0) Or _
            (adoRstLeaseDtl!InsurancePayable = "Y" And DateDiff("d", _
               IIf(adoRstLeaseDtl!InsuranceNextDueDate = "", _
               DateAdd("d", 1, CutOffDate), adoRstLeaseDtl!InsuranceNextDueDate), _
               CutOffDate) >= 0 And Val(IIf(IsNull(adoRstLeaseDtl!InsuranceEachPeriod), 0, adoRstLeaseDtl!InsuranceEachPeriod)) > 0) Then
            lDemand = NextRef(adoConn, "DEMAND_REF")        'GET THE DEMAND ID FROM THE REF TABLE
            AddNewRef adoConn, "DEMAND_REF", lDemand         'SET THE NEXT DEMAND ID IN THE REF TABLE

            With adoRstDemandRec
               .AddNew
               .Fields("DemandID").Value = lDemand
               .Fields("BatchID").Value = lBatch
               .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields("Source").Value = 1
               .Fields("TransactionType").Value = 1            'Invoice (I assume that in automatic demand all are invoice
'*** Here my thinking is, all type of demands due date is on the same day
'*** If its not correct then i have to change the manual demands grid and the demand table & split table
               .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
               .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields("IsPrinted").Value = False
               .Fields("UPDATE_SAGE").Value = False
               .Update
            End With

            iChildId = 0
'************************************************************************************
'*********************        Interest Charge            ****************************
'************************************************************************************
'*** Insert the split records in the DemandSplitRecords table
            If adoRstLeaseDtl.Fields("InterestChargeable").Value = "Y" Then
               iChildId = iChildId + 1
               szDes = "Interest Charges For " & adoRstLeaseDtl!DaysAfterInterestPayable & _
                        " Days on " & adoRstLeaseDtl!InterestChargedOn

'  ***  Add new demand IN demand table.
               adoRstSplitDemand.AddNew
               adoRstSplitDemand!SplitID = iChildId
               adoRstSplitDemand!DemandId = lDemand
               adoRstSplitDemand!A_M = "A"
               adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstLeaseDtl!IntDemandType), 0, " # ")
               adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstLeaseDtl!IntDemandType), 0, " # ")
               adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstLeaseDtl!IntDemandType), 2, " # ")
               adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstLeaseDtl!IntDemandType), 2, " # ")
               adoRstSplitDemand!Amount = adoRstLeaseDtl!InterestAmount
               adoRstSplitDemand!VATAmount = 0
               adoRstSplitDemand!TotalAmount = CCur(adoRstSplitDemand!Amount) + _
                                               CCur(adoRstSplitDemand!VATAmount)
'************************* ITS NEED TO BE CONFIRMED *******************************************************
' Currently INTEREST CHARGE does not have any due date and it also does not describe how the interest will be
' calculated. Interest should be calculated on any unpaid demands.
               adoRstSplitDemand!SageRef = szaPrefix(adoRstLeaseDtl!IntDemandType) & Format(DateAdd("d", CInt(adoRstLeaseDtl!DaysAfterInterestPayable), Date), "dd/mm/yy")
               adoRstSplitDemand!DueDate = Format(DateAdd("d", CInt(adoRstLeaseDtl!DaysAfterInterestPayable), Date), "dd/mm/yyyy")
'********************************************************************************************************
               adoRstSplitDemand!VATMonth = Month(Date)
               adoRstSplitDemand!typeofdemand = adoRstLeaseDtl!IntDemandType
               adoRstSplitDemand!description = szDes
               adoRstSplitDemand!DemandStatement = True
               adoRstSplitDemand.Update

               adoRstLeaseDtl!InterestChargeable = "N"
               adoRstLeaseDtl!DaysAfterInterestPayable = Null
               adoRstLeaseDtl!InterestChargedOn = Null
               adoRstLeaseDtl!InterestAmount = Null
               IPcount = IPcount + 1
               iSerial = iSerial + 1
            End If
'*********************************************************************************************************
'         Base Rent Demands
'*********************************************************************************************************
            szSQLStr = "SELECT * " & _
                       "FROM LRentCharges " & _
                       "WHERE LeaseID = '" & adoRstLeaseDtl!LeaseID & "';"

            Set adoRstRentCharges = New ADODB.Recordset
            adoRstRentCharges.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
            While Not adoRstRentCharges.EOF
               If adoRstLeaseDtl!BRPayable = "Y" And _
                  DateDiff("d", Date, IIf(adoRstRentCharges!BRNextDueDate = "", _
                     DateAdd("d", -1, Date), adoRstRentCharges!BRNextDueDate)) >= 0 And _
                  DateDiff("d", Date, IIf(adoRstRentCharges!BRNextDueDate = "", _
                     DateAdd("d", -1, Date), adoRstRentCharges!BRNextDueDate)) <= DaysB4Due Then
                  iChildId = iChildId + 1
      '**** Insert the Header info in the DemandRecords table
                  dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
                  szDes = DemandTypeName(adoRstRentCharges!BRDemandType, adoConn)
                  dtNtDueDate = FindNextDueDate(dtEndDate, adoRstRentCharges!BRNextDueDate, _
                                    adoRstRentCharges!BRfrequency)
                  adoRstSplitDemand.AddNew
                  adoRstSplitDemand!SplitID = iChildId
                  adoRstSplitDemand!DemandId = lDemand
                  adoRstSplitDemand!A_M = "A"
                  adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstRentCharges!BRDemandType), 0, " # ")
                  adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstRentCharges!BRDemandType), 0, " # ")
                  adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstRentCharges!BRDemandType), 1, " # ")
                  adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstRentCharges!BRDemandType), 1, " # ")
                  adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstRentCharges!BRDemandType), 2, " # ")
                  adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstRentCharges!BRDemandType), 2, " # ")
                  adoRstSplitDemand!Amount = adoRstRentCharges!BRAmount
                  adoRstSplitDemand!VATAmount = Round(adoRstRentCharges!BRAmount * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber) / 100, 2)
                  adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
                  adoRstSplitDemand!SageRef = szaPrefix(adoRstRentCharges!BRDemandType) & Format(adoRstRentCharges!BRNextDueDate, "dd/mm/yy")
                  adoRstSplitDemand!DueDate = Format(adoRstRentCharges!BRNextDueDate.Value, "dd/mm/yyyy")
                  adoRstSplitDemand!VATMonth = Month(adoRstRentCharges!BRNextDueDate)
                  adoRstSplitDemand!typeofdemand = adoRstRentCharges!BRDemandType
                  adoRstSplitDemand!description = szDes
                  adoRstSplitDemand!DemandStatement = True
                  adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber)
                  adoRstSplitDemand!DateFrom = CDate(adoRstRentCharges!BRNextDueDate)
                  adoRstSplitDemand!DateTo = dtNtDueDate
                  adoRstSplitDemand!SageDepartment = DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Rent Charges")
                  adoRstSplitDemand.Update

                  adoRstRentCharges!BRNextDueDate = dtNtDueDate
                  adoRstRentCharges.Update
                  BRcount = BRcount + 1
                  iSerial = iSerial + 1
               End If
               adoRstRentCharges.MoveNext
            Wend
            adoRstRentCharges.Close
            Set adoRstRentCharges = Nothing
'*********************************************************************************************************
'   Service Charge demands
'*********************************************************************************************************
            If adoRstLeaseDtl!SCPayable = "Y" And _
               DateDiff("d", Date, IIf(adoRstLeaseDtl!SCNextDueDate = "", _
                  DateAdd("d", -1, Date), adoRstLeaseDtl!SCNextDueDate)) >= 0 And _
               DateDiff("d", Date, IIf(adoRstLeaseDtl!SCNextDueDate = "", _
                  DateAdd("d", -1, Date), adoRstLeaseDtl!SCNextDueDate)) <= DaysB4Due Then
               iChildId = iChildId + 1
               dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
               szDes = "Service Charge" ' For Period " & adoRstLeaseDtl!SCNextDueDate & " to "
               dtNtDueDate = FindNextDueDate(dtEndDate, adoRstLeaseDtl!SCNextDueDate, _
                              adoRstLeaseDtl!SCfrequency)

               adoRstSplitDemand.AddNew
               adoRstSplitDemand!SplitID = iChildId
               adoRstSplitDemand!DemandId = lDemand
               adoRstSplitDemand!A_M = "A"
               adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstLeaseDtl!SCDemandType), 0, " # ")
               adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstLeaseDtl!SCDemandType), 0, " # ")
               adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstLeaseDtl!SCDemandType), 1, " # ")
               adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstLeaseDtl!SCDemandType), 1, " # ")
               adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstLeaseDtl!SCDemandType), 2, " # ")
               adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstLeaseDtl!SCDemandType), 2, " # ")
               adoRstSplitDemand!Amount = adoRstLeaseDtl!SCAmount
               adoRstSplitDemand!VATAmount = Round(adoRstLeaseDtl!SCAmount * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber) / 100, 2)
               adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
               adoRstSplitDemand!SageRef = szaPrefix(adoRstLeaseDtl!SCDemandType) & Format(adoRstLeaseDtl!SCNextDueDate, "dd/mm/yy")
               adoRstSplitDemand!DueDate = Format(adoRstLeaseDtl!SCNextDueDate.Value, "dd/mm/yyyy")
               adoRstSplitDemand!VATMonth = Month(adoRstLeaseDtl!SCNextDueDate)
               adoRstSplitDemand!typeofdemand = adoRstLeaseDtl!SCDemandType
               adoRstSplitDemand!description = szDes
               adoRstSplitDemand!DemandStatement = True
               adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber)
               adoRstSplitDemand!DateFrom = CDate(adoRstLeaseDtl!SCNextDueDate)
               adoRstSplitDemand!DateTo = dtNtDueDate
               adoRstSplitDemand!SageDepartment = DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Service Charge")
               adoRstSplitDemand.Update

               adoRstLeaseDtl!SCNextDueDate = dtNtDueDate     'Set to Next Due date in LeaseDetails table
               adoRstLeaseDtl.Update
               SCcount = SCcount + 1
               iSerial = iSerial + 1
            End If
         
'*********************************************************************************************************
'   Insurance Charge demands
'*********************************************************************************************************
            If adoRstLeaseDtl!InsurancePayable = "Y" And _
               DateDiff("d", Date, IIf(adoRstLeaseDtl!InsuranceNextDueDate = "", _
                  DateAdd("d", -1, Date), adoRstLeaseDtl!InsuranceNextDueDate)) >= 0 And _
               DateDiff("d", Date, IIf(adoRstLeaseDtl!InsuranceNextDueDate = "", _
                  DateAdd("d", -1, Date), adoRstLeaseDtl!InsuranceNextDueDate)) <= DaysB4Due Then
               iChildId = iChildId + 1
               dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
               szDes = "Insurance Charge" ' For Period " & adoRstLeaseDtl!SCNextDueDate & " to "
               dtNtDueDate = FindNextDueDate(dtEndDate, adoRstLeaseDtl!InsuranceNextDueDate, _
                              adoRstLeaseDtl!InsuranceFrequency)

               adoRstSplitDemand.AddNew
               adoRstSplitDemand!SplitID = iChildId
               adoRstSplitDemand!DemandId = lDemand
               adoRstSplitDemand!A_M = "A"
               adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstLeaseDtl!InsuranceDemandType), 0, " # ")
               adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstLeaseDtl!InsuranceDemandType), 0, " # ")
               adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstLeaseDtl!InsuranceDemandType), 1, " # ")
               adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstLeaseDtl!InsuranceDemandType), 1, " # ")
               adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstLeaseDtl!InsuranceDemandType), 2, " # ")
               adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstLeaseDtl!InsuranceDemandType), 2, " # ")
               adoRstSplitDemand!Amount = adoRstLeaseDtl!InsuranceEachPeriod
               adoRstSplitDemand!VATAmount = Round(adoRstLeaseDtl!InsuranceEachPeriod * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber) / 100, 2)
               adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
               adoRstSplitDemand!SageRef = szaPrefix(adoRstLeaseDtl!InsuranceDemandType) & Format(adoRstLeaseDtl!InsuranceNextDueDate, "dd/mm/yy")
               adoRstSplitDemand!DueDate = Format(adoRstLeaseDtl!InsuranceNextDueDate, "dd/mm/yyyy")
               adoRstSplitDemand!VATMonth = Month(adoRstLeaseDtl!InsuranceNextDueDate)
               adoRstSplitDemand!typeofdemand = adoRstLeaseDtl!InsuranceDemandType
               adoRstSplitDemand!description = szDes
               adoRstSplitDemand!DemandStatement = True
               adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber)
               adoRstSplitDemand!DateFrom = CDate(adoRstLeaseDtl!InsuranceNextDueDate)
               adoRstSplitDemand!DateTo = dtNtDueDate
               adoRstSplitDemand!SageDepartment = DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Insurance Charge")
               adoRstSplitDemand.Update

               adoRstLeaseDtl!InsuranceNextDueDate = dtNtDueDate     'Set to Next Due date in LeaseDetails table
               adoRstLeaseDtl.Update
               ICcount = ICcount + 1
               iSerial = iSerial + 1
            End If
         End If
         adoRstLeaseDtl.MoveNext
      Wend

      AddNewRef adoConn, "BATCH_REF", lBatch

      adoRstLeaseDtl.Close
      adoRstDemandRec.Close
      adoRstSplitDemand.Close
'
      Set adoRstLeaseDtl = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstSplitDemand = Nothing
   End If
'
   MousePointer = vbDefault
'
   Dim Msg As String
'
   Msg = Msg & BRcount & " Demands for Rent were generated." & Chr(13)
   Msg = Msg & SCcount & " Demands for Service Charge were generated." & Chr(13)
   Msg = Msg & IPcount & " Demands for Interest Payment were generated." & Chr(13)
   Msg = Msg & ICcount & " Demands for Insurance Charge were generated." & Chr(13)
   Msg = Msg & "A total of " & BRcount + SCcount + IPcount + ICcount & " demands were generated."

   MsgBox Msg, vbOKOnly + vbInformation, "Demands Generated"
   Exit Sub

ErrH:
       'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
       MsgBox ERR.Number & "-(Auto Con Demand) " & ERR.description, vbOKOnly, "Error"
       Resume Next
End Sub

Private Sub GenAutoSngDemands()
   Dim szaNCode() As String, szaNCName() As String, szaPrefix() As String
   Dim szDes As String, CutOffDate As String, szSQLStr As String

   Dim BRcount As Integer, SCcount As Integer, IPcount As Integer, ICcount As Integer
   Dim NextUniqueRefNo As Long, a As Integer, iSerial As Integer
   Dim lBatch As Long, lDemand As Long, iChildId As Integer

   Dim dtEndDate As Date, dtNtDueDate As Date

   Dim adoRstDemandRec As ADODB.Recordset, adoRstDmdTyp As ADODB.Recordset
   Dim adoRstLeaseDtl As ADODB.Recordset, adoRstSplitDemand As ADODB.Recordset
   Dim adoRstRentCharges As ADODB.Recordset
   Dim adoConn As ADODB.Connection
   Dim DaysB4Due As Integer

   If MsgBox("Are you sure to generate autometic demand?" & (Chr(13) + Chr(10)) & _
             "Make sure global data has been inputted correctly.", vbYesNo + vbQuestion, _
             "Generate Automatic Demands") = vbNo Then Exit Sub

   On Error GoTo ErrH

   MousePointer = vbHourglass    'change the mouse cursor to show program is busy/working

   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

'   Connect to Demands table to add new demands.
   Set adoRstDemandRec = New ADODB.Recordset
   szSQLStr = "SELECT * FROM DemandRecords"
   adoRstDemandRec.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   Set adoRstSplitDemand = New ADODB.Recordset
   szSQLStr = "SELECT * FROM DemandSplitRecords"
   adoRstSplitDemand.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

'   get nominal codes and prefix for base rent from demand types.
   Set adoRstDmdTyp = New ADODB.Recordset
   szSQLStr = "SELECT * FROM DemandTypes;"
   adoRstDmdTyp.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly

   If adoRstDmdTyp.EOF Then
      MsgBox "There are no Demand Type in the database", vbCritical, "DATA Input Error"

      adoRstSplitDemand.Close
      adoRstDemandRec.Close
      adoRstDmdTyp.Close
      adoConn.Close
      Set adoRstSplitDemand = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstDmdTyp = Nothing
      Set adoConn = Nothing
      Exit Sub
   End If
   ReDim szaPrefix(adoRstDmdTyp.RecordCount) As String
   ReDim szaNCName(adoRstDmdTyp.RecordCount) As String
   ReDim szaNCode(adoRstDmdTyp.RecordCount) As String
'
'**Saving all Nominal Codes and Prefixes in the array
   While Not adoRstDmdTyp.EOF
      szaPrefix(adoRstDmdTyp!ID) = adoRstDmdTyp!Prefix
      szaNCName(adoRstDmdTyp!ID) = adoRstDmdTyp!NominalNameforAmount & " # " & adoRstDmdTyp!NominalNameforVAT & " # " & adoRstDmdTyp!NominalNameforTotal
      szaNCode(adoRstDmdTyp!ID) = adoRstDmdTyp!NominalCodeforAmount & " # " & adoRstDmdTyp!NominalCodeForVAT & " # " & adoRstDmdTyp!NominalCodeForTotal
      adoRstDmdTyp.MoveNext
   Wend
   adoRstDmdTyp.Close
   Set adoRstDmdTyp = Nothing

   Set adoRstLeaseDtl = New ADODB.Recordset

   szSQLStr = "SELECT * FROM LeaseDetails WHERE Status = TRUE;"
   adoRstLeaseDtl.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   iSerial = 1
   lBatch = NextRef(adoConn, "BATCH_REF")
   If adoRstLeaseDtl.EOF Then
       adoRstLeaseDtl.Close
       Set adoRstLeaseDtl = Nothing
   Else
      While Not adoRstLeaseDtl.EOF

'*********************** SAMRAT 25/11/2005***************************************
'*** Determin the date boundray in future, in this boundary
'*** we have to find those demands' due date to calcuate
'*** demands from lease table and we have to calculate &
'*** collect all those demands in DemandRecords table.
'*********************************************************************************
         DaysB4Due = GlbDaysBeforeDue(adoRstLeaseDtl.Fields("LeaseID").Value)
         CutOffDate = DateAdd("d", DaysB4Due, Date) 'Date boundary
'************************************************************************************
'*********************        Interest Charge            ****************************
'************************************************************************************
         If adoRstLeaseDtl.Fields("InterestChargeable").Value = "Y" Then
'**** Insert the Header info in the DemandRecords table
            lDemand = NextRef(adoConn, "DEMAND_REF")        'GET THE DEMAND ID FROM THE REF TABLE
            AddNewRef adoConn, "DEMAND_REF", lDemand        'SET THE NEXT DEMAND ID IN THE REF TABLE

            With adoRstDemandRec
               .AddNew
               .Fields("DemandID").Value = lDemand
               .Fields("BatchID").Value = lBatch
               .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields("Source").Value = 1
               .Fields("TransactionType").Value = 1
'*** Here my thinking is, all type of demands due date is on the same day
'*** If its not correct then i have to change the manual demands grid and the demand table & split table
               .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
               .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields("IsPrinted").Value = False
               .Fields("UPDATE_SAGE").Value = False
               .Update
            End With
            
'*** Insert the split records in the DemandSplitRecords table
            iChildId = 1
            szDes = "Interest Charges For " & adoRstLeaseDtl!DaysAfterInterestPayable & _
                     " Days on " & adoRstLeaseDtl!InterestChargedOn

'*** Add new demand IN demand table.
            adoRstSplitDemand.AddNew
            adoRstSplitDemand!SplitID = iChildId
            adoRstSplitDemand!DemandId = lDemand
            adoRstSplitDemand!A_M = "A"
            adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstLeaseDtl!IntDemandType), 0, " # ")
            adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstLeaseDtl!IntDemandType), 0, " # ")
            adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstLeaseDtl!IntDemandType), 2, " # ")
            adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstLeaseDtl!IntDemandType), 2, " # ")
            adoRstSplitDemand!Amount = adoRstLeaseDtl!InterestAmount
            adoRstSplitDemand!VATAmount = 0
            adoRstSplitDemand!TotalAmount = CCur(adoRstSplitDemand!Amount) + _
                                            CCur(adoRstSplitDemand!VATAmount)
'************************* ITS NEED TO BE CONFIRMED *******************************************************
' Currently INTEREST CHARGE does not have any due date and it also does not describe how the interest will be
' calculated. Interest should be calculated on any unpaid demands.
            adoRstSplitDemand!SageRef = szaPrefix(adoRstLeaseDtl!IntDemandType) & Format(DateAdd("d", CInt(adoRstLeaseDtl!DaysAfterInterestPayable), Date), "dd/mm/yy")
            adoRstSplitDemand!DueDate = Format(DateAdd("d", CInt(adoRstLeaseDtl!DaysAfterInterestPayable), Date), "dd/mm/yyyy")
'********************************************************************************************************
            adoRstSplitDemand!VATMonth = Month(Date)
            adoRstSplitDemand!typeofdemand = adoRstLeaseDtl!IntDemandType
            adoRstSplitDemand!description = szDes
            adoRstSplitDemand!DemandStatement = True
            adoRstSplitDemand.Update

            adoRstLeaseDtl!InterestChargeable = "N"
            adoRstLeaseDtl!DaysAfterInterestPayable = Null
            adoRstLeaseDtl!InterestChargedOn = Null
            adoRstLeaseDtl!InterestAmount = Null
            IPcount = IPcount + 1
            iSerial = iSerial + 1
         End If
'*********************************************************************************************************
'         Base Rent Demands
'*********************************************************************************************************
         szSQLStr = "SELECT * " & _
                    "FROM LRentCharges " & _
                    "WHERE LeaseID = '" & adoRstLeaseDtl!LeaseID & "';"

         Set adoRstRentCharges = New ADODB.Recordset
         adoRstRentCharges.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
         While Not adoRstRentCharges.EOF
            If adoRstLeaseDtl!BRPayable = "Y" And _
               DateDiff("d", Date, IIf(adoRstRentCharges!BRNextDueDate = "", _
                  DateAdd("d", -1, Date), adoRstRentCharges!BRNextDueDate)) >= 0 And _
               DateDiff("d", Date, IIf(adoRstRentCharges!BRNextDueDate = "", _
                  DateAdd("d", -1, Date), adoRstRentCharges!BRNextDueDate)) <= DaysB4Due Then
   '**** Insert the Header info in the DemandRecords table
               lDemand = NextRef(adoConn, "DEMAND_REF")        'GET THE DEMAND ID
               AddNewRef adoConn, "DEMAND_REF", lDemand
               With adoRstDemandRec
                  .AddNew
                  .Fields("DemandID").Value = lDemand
                  .Fields("BatchID").Value = lBatch
                  .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber.Value
                  .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName.Value
                  .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber.Value
                  .Fields("Source").Value = 1
                  .Fields("TransactionType").Value = 1
   '***Here my thinking is, all type of demands due date is on the same day
   '***If its not correct then i have to change the manual demands grid and the demand table & split table
                  .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
                  .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
                  .Fields("IsPrinted").Value = False
                  .Fields("UPDATE_SAGE").Value = False
                  .Update
               End With

               iChildId = 1
               dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
               szDes = "Rent Charges"
               dtNtDueDate = FindNextDueDate(dtEndDate, adoRstRentCharges!BRNextDueDate, _
                                 adoRstRentCharges!BRfrequency)
               adoRstSplitDemand.AddNew
               adoRstSplitDemand!SplitID = iChildId
               adoRstSplitDemand!DemandId = lDemand
               adoRstSplitDemand!A_M = "A"
               adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstRentCharges!BRDemandType), 0, " # ")
               adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstRentCharges!BRDemandType), 0, " # ")
               adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstRentCharges!BRDemandType), 1, " # ")
               adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstRentCharges!BRDemandType), 1, " # ")
               adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstRentCharges!BRDemandType), 2, " # ")
               adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstRentCharges!BRDemandType), 2, " # ")
               adoRstSplitDemand!Amount = adoRstRentCharges!BRAmount
               adoRstSplitDemand!VATAmount = Round(adoRstRentCharges!BRAmount * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber) / 100, 2)
               adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
               adoRstSplitDemand!SageRef = szaPrefix(adoRstRentCharges!BRDemandType) & Format(adoRstRentCharges!BRNextDueDate, "dd/mm/yy")
               adoRstSplitDemand!DueDate = Format(adoRstRentCharges!BRNextDueDate.Value, "dd/mm/yyyy")
               adoRstSplitDemand!VATMonth = Month(adoRstRentCharges!BRNextDueDate)
               adoRstSplitDemand!typeofdemand = adoRstRentCharges!BRDemandType
               adoRstSplitDemand!description = szDes
               adoRstSplitDemand!DemandStatement = True
               adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber)
               adoRstSplitDemand!DateFrom = CDate(adoRstRentCharges!BRNextDueDate)
               adoRstSplitDemand!DateTo = dtNtDueDate
               adoRstSplitDemand!SageDepartment = DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Rent Charges")
               adoRstSplitDemand.Update

               adoRstRentCharges!BRNextDueDate = dtNtDueDate
               adoRstRentCharges.Update
               BRcount = BRcount + 1
               iSerial = iSerial + 1
            End If
            adoRstRentCharges.MoveNext
         Wend
         adoRstRentCharges.Close
         Set adoRstRentCharges = Nothing
'************************************************************************************************
'   Service Charge demands
'************************************************************************************************
         If adoRstLeaseDtl!SCPayable = "Y" And _
            DateDiff("d", Date, IIf(adoRstLeaseDtl!SCNextDueDate = "", _
               DateAdd("d", -1, Date), adoRstLeaseDtl!SCNextDueDate)) >= 0 And _
            DateDiff("d", Date, IIf(adoRstLeaseDtl!SCNextDueDate = "", _
               DateAdd("d", -1, Date), adoRstLeaseDtl!SCNextDueDate)) <= DaysB4Due Then
'**** Insert the Header info in the DemandRecords table
            lDemand = NextRef(adoConn, "DEMAND_REF")        'GET THE DEMAND ID
            AddNewRef adoConn, "DEMAND_REF", lDemand
            With adoRstDemandRec
               .AddNew
               .Fields("DemandID").Value = lDemand
               .Fields("BatchID").Value = lBatch
               .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields("Source").Value = 1
               .Fields("TransactionType").Value = 1
'***Here my thinking is, all type of demands due date is on the same day
'***If its not correct then i have to change the manual demands grid and the demand table & split table
               .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
               .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields("IsPrinted").Value = False
               .Fields("UPDATE_SAGE").Value = False
               .Update
            End With

            iChildId = 1
            dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
            szDes = "Service Charge" ' For Period " & adoRstLeaseDtl!SCNextDueDate & " to "
            dtNtDueDate = FindNextDueDate(dtEndDate, adoRstLeaseDtl!SCNextDueDate, _
                           adoRstLeaseDtl!SCfrequency)

            adoRstSplitDemand.AddNew
            adoRstSplitDemand!SplitID = iChildId
            adoRstSplitDemand!DemandId = lDemand
            adoRstSplitDemand!A_M = "A"
            adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstLeaseDtl!SCDemandType), 0, " # ")
            adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstLeaseDtl!SCDemandType), 0, " # ")
            adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstLeaseDtl!SCDemandType), 1, " # ")
            adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstLeaseDtl!SCDemandType), 1, " # ")
            adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstLeaseDtl!SCDemandType), 2, " # ")
            adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstLeaseDtl!SCDemandType), 2, " # ")
            adoRstSplitDemand!Amount = adoRstLeaseDtl!SCAmount
            adoRstSplitDemand!VATAmount = Round(adoRstLeaseDtl!SCAmount * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber) / 100, 2)
            adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
            adoRstSplitDemand!SageRef = szaPrefix(adoRstLeaseDtl!SCDemandType) & Format(adoRstLeaseDtl!SCNextDueDate, "dd/mm/yy")
            adoRstSplitDemand!DueDate = Format(adoRstLeaseDtl!SCNextDueDate.Value, "dd/mm/yyyy")
            adoRstSplitDemand!VATMonth = Month(adoRstLeaseDtl!SCNextDueDate)
            adoRstSplitDemand!typeofdemand = adoRstLeaseDtl!SCDemandType
            adoRstSplitDemand!description = szDes
            adoRstSplitDemand!DemandStatement = True
            adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber)
            adoRstSplitDemand!DateFrom = CDate(adoRstLeaseDtl!SCNextDueDate)
            adoRstSplitDemand!DateTo = dtNtDueDate
            adoRstSplitDemand!SageDepartment = DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Service Charge")
            adoRstSplitDemand.Update

            adoRstLeaseDtl!SCNextDueDate = dtNtDueDate     'Set to Next Due date in LeaseDetails table
            adoRstLeaseDtl.Update
            SCcount = SCcount + 1
            iSerial = iSerial + 1
         End If
'************************************************************************************************
'   Insurance Charge demands
'************************************************************************************************
         If adoRstLeaseDtl!InsurancePayable = "Y" And _
            DateDiff("d", Date, IIf(adoRstLeaseDtl!InsuranceNextDueDate = "", _
               DateAdd("d", -1, Date), adoRstLeaseDtl!InsuranceNextDueDate)) >= 0 And _
            DateDiff("d", Date, IIf(adoRstLeaseDtl!InsuranceNextDueDate = "", _
               DateAdd("d", -1, Date), adoRstLeaseDtl!InsuranceNextDueDate)) <= DaysB4Due Then
'**** Insert the Header info in the DemandRecords table
            lDemand = NextRef(adoConn, "DEMAND_REF")        'GET THE DEMAND ID
            AddNewRef adoConn, "DEMAND_REF", lDemand
            With adoRstDemandRec
               .AddNew
               .Fields("DemandID").Value = lDemand
               .Fields("BatchID").Value = lBatch
               .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields("Source").Value = 1
               .Fields("TransactionType").Value = 1
'***Here my thinking is, all type of demands due date is on the same day
'***If its not correct then i have to change the manual demands grid and the demand table & split table
               .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
               .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields("IsPrinted").Value = False
               .Fields("UPDATE_SAGE").Value = False
               .Update
            End With

            iChildId = 1
            dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
            szDes = "Insurance Charge"
            dtNtDueDate = FindNextDueDate(dtEndDate, adoRstLeaseDtl!InsuranceNextDueDate, _
                           adoRstLeaseDtl!InsuranceFrequency)

            adoRstSplitDemand.AddNew
            adoRstSplitDemand!SplitID = iChildId
            adoRstSplitDemand!DemandId = lDemand
            adoRstSplitDemand!A_M = "A"
            adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstLeaseDtl!InsuranceDemandType), 0, " # ")
            adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstLeaseDtl!InsuranceDemandType), 0, " # ")
            adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstLeaseDtl!InsuranceDemandType), 1, " # ")
            adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstLeaseDtl!InsuranceDemandType), 1, " # ")
            adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstLeaseDtl!InsuranceDemandType), 2, " # ")
            adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstLeaseDtl!InsuranceDemandType), 2, " # ")
            adoRstSplitDemand!Amount = adoRstLeaseDtl!InsuranceEachPeriod
            adoRstSplitDemand!VATAmount = Round(adoRstLeaseDtl!InsuranceEachPeriod * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber) / 100, 2)
            adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
            adoRstSplitDemand!SageRef = szaPrefix(adoRstLeaseDtl!InsuranceDemandType) & Format(adoRstLeaseDtl!InsuranceNextDueDate, "dd/mm/yy")
            adoRstSplitDemand!DueDate = Format(adoRstLeaseDtl!InsuranceNextDueDate, "dd/mm/yyyy")
            adoRstSplitDemand!VATMonth = Month(adoRstLeaseDtl!InsuranceNextDueDate)
            adoRstSplitDemand!typeofdemand = adoRstLeaseDtl!InsuranceDemandType
            adoRstSplitDemand!description = szDes
            adoRstSplitDemand!DemandStatement = True
            adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber)
            adoRstSplitDemand!DateFrom = CDate(adoRstLeaseDtl!InsuranceNextDueDate)
            adoRstSplitDemand!DateTo = dtNtDueDate
            adoRstSplitDemand!SageDepartment = DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Insurance Charge")
            adoRstSplitDemand.Update

            adoRstLeaseDtl!InsuranceNextDueDate = dtNtDueDate     'Set to Next Due date in LeaseDetails table
            adoRstLeaseDtl.Update
            ICcount = ICcount + 1
            iSerial = iSerial + 1
         End If
         adoRstLeaseDtl.MoveNext
      Wend

      AddNewRef adoConn, "BATCH_REF", lBatch
      adoRstLeaseDtl.Close
      adoRstDemandRec.Close
      adoRstSplitDemand.Close

      Set adoRstLeaseDtl = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstSplitDemand = Nothing
   End If

   MousePointer = vbDefault

   Dim Msg As String

   Msg = Msg & BRcount & " Demands for Rent were generated." & Chr(13)
   Msg = Msg & SCcount & " Demands for Service Charge were generated." & Chr(13)
   Msg = Msg & IPcount & " Demands for Interest Payment were generated." & Chr(13)
   Msg = Msg & ICcount & " Demands for Insurance Payment were generated." & Chr(13)
   Msg = Msg & "A total of " & BRcount + SCcount + IPcount + ICcount & " demands were generated."

   MsgBox Msg, vbOKOnly + vbInformation, "Demands Generated"
   Exit Sub
ErrH:
       'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
       MsgBox ERR.Number & " - (pcm_001)" & ERR.description, vbOKOnly, "Error"
       Resume Next
End Sub

Private Function GeneratableBaseRent(ByVal adoConn As ADODB.Connection, ByVal szLeaseID As String, DaysB4Due As Integer) As Boolean
   Dim szSQLStr As String
   Dim adoRstRentCharges As ADODB.Recordset

   GeneratableBaseRent = False

   szSQLStr = "SELECT * " & _
              "FROM LRentCharges " & _
              "WHERE LeaseID = '" & szLeaseID & "';"

   Set adoRstRentCharges = New ADODB.Recordset
   adoRstRentCharges.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
   While Not adoRstRentCharges.EOF
      If DateDiff("d", Date, IIf(adoRstRentCharges!BRNextDueDate = "", _
            DateAdd("d", -1, Date), adoRstRentCharges!BRNextDueDate)) >= 0 And _
         DateDiff("d", Date, IIf(adoRstRentCharges!BRNextDueDate = "", _
            DateAdd("d", -1, Date), adoRstRentCharges!BRNextDueDate)) <= DaysB4Due And _
            Val(IIf(IsNull(adoRstRentCharges!BRAmount), 0, adoRstRentCharges!BRAmount)) > 0 Then
         GeneratableBaseRent = True
      End If
      adoRstRentCharges.MoveNext
   Wend
   adoRstRentCharges.Close
   Set adoRstRentCharges = Nothing
End Function

Private Function FindNextDueDate(dtEndDate As Date, dtNtDueDate As Date, iFreq As Integer) As Date
   Dim dtCurDueDt As Date

   If CInt(Left(szaFreq(iFreq), Len(szaFreq(iFreq)) - 1)) > 0 Then      'In advance
     FindNextDueDate = DateAdd(Right(szaFreq(iFreq), 1), CDbl(Left(szaFreq(iFreq), _
                 Len(szaFreq(iFreq)) - 1)), CDate(dtNtDueDate))           ' interval, number, date
      If DateDiff("d", dtEndDate, FindNextDueDate) > 0 Then
         FindNextDueDate = dtEndDate
      End If
   Else                                                                  'In arrears
      FindNextDueDate = DateAdd("d", -1, CDate(dtNtDueDate))
      If DateDiff("d", dtEndDate, FindNextDueDate) > 0 Then
         FindNextDueDate = dtEndDate
      End If
   End If
End Function

Private Sub ShowReport(szReportFileName As String)
       'Declare the application object used to open the rpt file
   Dim crxApplication As New CRAXDRT.Application
'
   'Declare the report object
   Dim Report As CRAXDRT.Report
   Dim i As Integer
'Debug.Print szReportFileName
'   Set Report = crxApplication.OpenReport(App.Path & "\" & szReportFileName, 1)
   Set Report = crxApplication.OpenReport(szReportFileName, 1)

   Dim strInvoiceAmt As String
   Dim rep As frmReport

   Set rep = New frmReport

'   For i = 1 To Report.Database.Tables.Count
'       Report.Database.Tables(i).SetLogOnInfo szPictureDBPath, szPictureDBPath, , ""
'   Next i
   
   rep.LoadReportViewer Report
'   rep.Refresh
End Sub

Public Sub DisableMenu()

'mnuGenManual.Enabled = False
'mnuDeleteDemand.Enabled = False
'mnuEditDemands.Enabled = False
'mnuPrint.Enabled = False
'mnuReprint.Enabled = False
'mnuPrintBatch.Enabled = False

End Sub

Public Sub EnableMenu()
'mnuGenManual.Enabled = True
'mnuDeleteDemand.Enabled = True
'mnuEditDemands.Enabled = True
'mnuPrint.Enabled = True
'mnuReprint.Enabled = True
'mnuPrintBatch.Enabled = True
End Sub

Public Sub ReprintDemands()

'ReprintMode = True
'
'Set adoConn = New adodb.Connection
'Set adoRst = New adodb.Recordset
'
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE IsPrinted = 'Y'"
'adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
'If adoRst.EOF = False Then
'    While adoRst.EOF = False
'        adoRst!SendToPrint = "Y"
'        adoRst.Update
'        adoRst.MoveNext
'    Wend
'Else
'    MsgBox "There are no demands to reprint.", vbOKOnly + vbInformation, "Reprint"
'    adoRst.Close
'    adoConn.Close
'    Set adoRst = Nothing
'    Set adoConn = Nothing
'    ReprintMode = False
'    Exit Sub
'End If
'
'adoRst.Close
'adoConn.Close
'Set adoRst = Nothing
'Set adoConn = Nothing
'
'Call EmptyBoxes
'Call GetFirstDemand
'
'chkPrint.Visible = True
'cmdGenAll.Visible = False
'cmdGenerateManual.Visible = False
'cmdEdit.Visible = False
'cmdDelete.Visible = False
'cmdDeleteOld.Visible = False
'cmdReprint.Visible = False
'cmdPrint.Visible = False
'cmdPrintThis.Visible = False
'cmdReprintAll.Visible = True
'cmdPrintBatch.Visible = False
'cmdReprintSome.Visible = True
'cmdCancelReprint.Visible = True
''lbl1.Visible = True
'
'Call DisableMenu

End Sub

Public Sub PrintBatch(BatchToPrint As Integer)

'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

'rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
'rdoConn.CursorDriver = rdUseIfNeeded
'rdoConn.EstablishConnection rdDriverNoPrompt
'
'SQLStr1 = "SELECT * FROM Batches WHERE BATCHID = " & BatchToPrint
'Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'If rdoRst1.EOF = True And rdoRst1.EOF = True Then ' no record
'    MsgBox "Invalid BATCHID number entered.", "Invalid BATCHID"
'    Exit Sub
'End If
'
'If MsgBox("Print BATCHID " & BatchToPrint & ": Demands " & rdoRst1!first & " to " & rdoRst1!last & " ?", vbYesNo, "Print BATCHID") = vbNo Then Exit Sub
'
'rdoRst1.Close
'
'SQLStr1 = "SELECT UniqueRefNumber, IsPrinted FROM DemandRecords WHERE BATCHID = " & BatchToPrint
'Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
'
'If rdoRst1.EOF = False Then
'    While rdoRst1.EOF = False
'        rdoRst1.Edit
'        'MsgBox rdoRst1!UniqueRefNumber
'        rdoRst1!IsPrinted = "C"
'        rdoRst1.Update
'        rdoRst1.MoveNext
'    Wend
'End If
'rdoRst1.Close
'
''send demands to crystal.
'CR1.ReportFileName = App.Path & "\REPORTFILES\Demand" & SCID & ".rpt"
'MsgBox CR1.ReportFileName
'CR1.printReport
'
'Call SetPrintedtoYes

End Sub


'Public Sub GetUniqueReferenceNumber()

'Dim Source As String
'Dim a As Integer
'Set adoConn = New ADODB.Connection
'Set adoRst = New ADODB.Recordset

'adoConn.Open "DSN=" & Adsn & ";UID=;PWD=;"

'Source = "SELECT UniqueRefNumber FROM DemandRecords"
'adoRst.Open Source, adoConn, adOpenStatic, adLockReadOnly

'If adoRst.EOF = False Then
    'While adoRst.EOF = False
     '   If a < adoRst!UniqueRefNumber Then a = adoRst!UniqueRefNumber
    '    adoRst.MoveNext
   ' Wend
  '  adoRst.MoveFirst
 '   Call GetRecord
'End If

'NextUniqueRefNo = a + 1
'adoRst.Close
'adoConn.Close
'Set adoConn = Nothing

'End Sub

Public Sub GetFirstDemand()
'    Dim a As Integer
'
'    Set adoConn = New adodb.Connection
'    Set adoRst = New adodb.Recordset
'
'    adoConn.Open "DSN=" & Adsn & ";UID=;PWD=;"
'
'    strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords"
'    If EditMode Then _
'         strSQLTitles = "SELECT UniqueRefNumber " & _
'                        "FROM DemandRecords  " & _
'                        "WHERE UPDATE_SAGE = 'N'"
'    If PrintMode Then _
'         strSQLTitles = "SELECT UniqueRefNumber  " & _
'                        "FROM DemandRecords  " & _
'                        "WHERE IsPrinted = 'N'"
'    If ReprintMode Then _
'         strSQLTitles = "SELECT UniqueRefNumber  " & _
'                        "FROM DemandRecords  " & _
'                        "WHERE IsPrinted = 'Y'"
'
'    adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
'
'    If Not adoRst.EOF Then
'        a = adoRst!UniqueRefNumber
'
'        While adoRst.EOF = False
'            If a > adoRst!UniqueRefNumber Then a = adoRst!UniqueRefNumber
'
'            adoRst.MoveNext
'        Wend
'        adoRst.Close
'        Set adoRst = Nothing
'        Set adoRst = New adodb.Recordset
'        strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & a & "';"
'        adoRst.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly
'        Call GetRecord
''        cmdMoveFirst.Visible = True
''        cmdMoveNext.Visible = True
''        cmdMoveLast.Visible = True
''        cmdMovePrevious.Visible = True
'        adoRst.Close
'        adoConn.Close
'        Set adoRst = Nothing
'        Set adoConn = Nothing
'    Else
'        MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
'        Exit Sub
'    End If
End Sub

Public Sub UpdatePrint()

'Set adoConn = New ADODB.Connection
'Set adoRst = New ADODB.Recordset
'
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE UniqueRefNumber = '" & Text1.text & "';"
'adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
'If chkPrint.Value = 0 Then ' user has selected to not print the current demand
'    adoRst!SendToPrint = ""
'    adoRst.Update
'Else
'    adoRst!SendToPrint = "Y"
'    adoRst.Update
'End If
'
'adoRst.Close
'adoConn.Close
'Set adoRst = Nothing
'Set adoConn = Nothing

End Sub

Public Sub UnSelect()

'Set adoConn = New adodb.Connection
'Set adoRst = New adodb.Recordset
'
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
'adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
'If adoRst.EOF = False Then
'    While adoRst.EOF = False
'        adoRst!SendToPrint = ""
'        adoRst.Update
'        adoRst.MoveNext
'    Wend
'End If
'adoRst.Close
'adoConn.Close
'Set adoConn = Nothing

End Sub

Public Sub SetPrintedtoYes()

'Set adoConn = New adodb.Connection
'Set adoRst = New adodb.Recordset
'
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'
'strSQLTitles = "SELECT IsPrinted FROM DemandRecords WHERE IsPrinted = 'C'"
'adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'
'If adoRst.EOF = False Then
'    While adoRst.EOF = False
'        adoRst!IsPrinted = "Y"
'        adoRst.Update
'        adoRst.MoveNext
'        DoEvents
'    Wend
'End If
'
'adoRst.Close
'adoConn.Close
'Set adoConn = Nothing

End Sub

Public Sub PrintBatchSelected()

'to go to timeout
'Call CheckDateAndTimeoutFileNoKey

'   Dim Response
'   Dim batchnum As String
'   Dim match As Boolean
'   Dim NumOfBatches As Integer
'   Dim i As Integer
'   match = False
'
'   rdoConn.Connect = "DSN=" & Adsn & ";UID=PWD=;"
'   rdoConn.CursorDriver = rdUseIfNeeded
'   rdoConn.EstablishConnection rdDriverNoPrompt
'
'   SQLStr1 = "SELECT BATCHID FROM batches"
'   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'   If rdoRst1.EOF = False Then
'       rdoRst1.MoveLast
'       rdoRst1.MoveFirst
'       NumOfBatches = rdoRst1.RowCount
'       ReDim GetBatches(NumOfBatches) As Integer
'       i = 1
'       While rdoRst1.EOF = False
'           GetBatches(i) = rdoRst1!BATCHID
'           rdoRst1.MoveNext
'           i = i + 1
'       Wend
'   End If
'   rdoRst1.Close
'   rdoConn.Close
'
'ReenterBatch:
'   batchnum = InputBox("Enter BATCHID to print: ", "Print BATCHID")
'
'   If IsNumeric(batchnum) = False Then
'       While IsNumeric(batchnum) = False
'           Response = MsgBox("You have entered an invalid BATCHID number.", vbRetryCancel, "Incorrect BATCHID")
'           If Response = vbCancel Then Exit Sub
'           If Response = vbRetry Then batchnum = InputBox("Enter a BATCHID to print: ", "Print BATCHID")
'       Wend
'   End If
'
'   For i = 1 To NumOfBatches
'       If GetBatches(i) = CInt(batchnum) Then match = True
'   Next i
'   If match = False Then 'not valid batchnumber
'       Response = MsgBox("You have entered an invalid BATCHID number.", vbRetryCancel, "Invalid BATCHID")
'       If Response = vbCancel Then Exit Sub
'       If Response = vbRetry Then GoTo ReenterBatch
'   Else 'valid BATCHID number
'       Call UnSelect
'       Call PrintBatch(CInt(batchnum))
'   End If

End Sub

Public Sub SaveChanges()

'If MsgBox("Save changes to current demand?", vbSystemModal + vbQuestion + vbYesNo, "Save Changes") = vbYes Then
'    Set adoConn = New ADODB.Connection
'    Set adoRst = New ADODB.Recordset
'    adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'    strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = '" & Text1.text & "';"
'    adoRst.Open strSQLTitles, adoConn, adOpenDynamic, adLockPessimistic
'    If szIC = "I" Then adoRst!TransactionType = 4
'    If szIC = "C" Then adoRst!TransactionType = 5
'    adoRst!TotalAmount = txtTotal.text
'    adoRst!VATAmount = txtVatAmt.text
'    adoRst!Amount = txtAmount.text
'    If txt4.text <> "" Then adoRst!IssueDate = Left(txt4.text, 6) + Right(txt4.text, 2)
'    adoRst!DueDate = Left(txtDueDate.text, 6) + Right(txtDueDate.text, 2)
'    adoRst!VATMonth = Month(txt4.text)
'    If cboType.ListIndex = -1 Then
'        adoRst!typeofdemand = typeofdemand
'    Else
'        adoRst!typeofdemand = cboType.ListIndex
'    End If
'    adoRst!Reference = txtRef.text
'    adoRst!text = txtSageText.text
'    adoRst!description = txtDescription.text
'    adoRst.Update
'    adoRst.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End If

End Sub


Public Sub CancelPrint(a As Integer)

'Set adoConn = New adodb.Connection
'Set adoRst = New adodb.Recordset
'
'adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'SQLStr1 = "SELECT IsPrinted FROM DemandRecords WHERE BATCHID = " & a
'adoRst.Open SQLStr1, adoConn, adOpenDynamic, adLockPessimistic
'
'While adoRst.EOF = False
'    adoRst!IsPrinted = "N"
'    adoRst.Update
'    adoRst.MoveNext
'Wend
'
'adoRst.Close
'adoConn.Close
'Set adoRst = Nothing
'Set adoConn = Nothing
'
End Sub

Public Sub DeleteDemands()
'   Set adoConn = New adodb.Connection
'   Set adoRst = New adodb.Recordset
'
'   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
'   SQLStr1 = "SELECT * FROM DemandRecords WHERE UPDATE_SAGE = 'Y' AND ExportedToExcel = 'Y'"
'   adoRst.Open SQLStr1, adoConn, adOpenDynamic, adLockPessimistic
'
'   If adoRst.EOF = False Then
'       While adoRst.EOF = False
'           adoRst.Delete
'           adoRst.MoveNext
'       Wend
'   End If
'
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
End Sub

Private Sub cmdClearDemands_Click()
   MousePointer = vbHourglass
   
   'user wants to delete all old demands - ones that have been printed, exported to sage and exported to excel
   If MsgBox("Are you sure to clear down all current demand records? All demands will be permanently deleted?", vbYesNo + vbQuestion, "Delete Old Demands") = vbNo Then Exit Sub

   Call ClearTable("DEMANDRECORDS")

   Call GetFirstDemand
   MsgBox "Old demands deleted successfully", vbOKOnly + vbInformation, "Deleted"
   MousePointer = vbDefault
End Sub

Public Sub ClearTable(szTable As String)
   Dim adoConn As ADODB.Connection
   Dim adoRst As ADODB.Recordset
   Dim SQLStr1 As String

   Set adoConn = New ADODB.Connection
   Set adoRst = New ADODB.Recordset

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="

   SQLStr1 = "DELETE * FROM " & szTable & ";"
   adoRst.Open SQLStr1, adoConn, adOpenDynamic, adLockPessimistic

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub lblReprintDemand_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblReprintDemand.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
End Sub

Private Sub optBatchWise_Click()
   cmbBatchWise.Enabled = True
   cmbBatchWise.SetFocus
   fraChoice3.Enabled = True
End Sub

Private Sub optInvoice_Click()
   fraChoice2.Enabled = True
End Sub

Private Sub optPayeeWise_Click()
   cmbBatchWise.text = ""
   cmbBatchWise.Enabled = False
   fraChoice3.Enabled = True
End Sub

Private Sub optReport_Click()
   fraChoice2.Enabled = True
End Sub

Private Sub tabDmdRcpt_Click(PreviousTab As Integer)
   MousePointer = vbHourglass
   tabPayment.Tab = 0
   
   If tabDmdRcpt.Tab = 2 Then       'Receipt
      If cmbBankAc.ListCount < 1 Then
         szAllBankBalance = BankAndBalance(cmbBankAc)
      End If
      If cmbTenant.ListCount < 1 Then
         TeantComboBox cmbTenant
      End If
      txtSPDate.text = Format(Now, "dd/mm/yyyy")
      cmbBankAc.SetFocus
   End If
   
   MousePointer = vbDefault
End Sub

Private Sub LoadDataInGrid()
   Dim rdoConn As New RDO.rdoConnection
   Dim rdoRst1 As rdoResultset
   Dim SQLStr1 As String, szaTenant() As String
   Dim iRow As Integer
'
   szaTenant = Split(cmbTenant.text, " \ ")
'   connect to database
   rdoConn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   rdoConn.CursorDriver = rdUseIfNeeded
   rdoConn.EstablishConnection rdDriverNoPrompt
'
'   Get the details for the demand type selected
   SQLStr1 = "SELECT tlbReceipt.*, tlbTransactionTypes.DESCRIPTION " & _
             "FROM tlbReceipt, tlbTransactionTypes " & _
             "WHERE SageAccountNumber = '" & szaTenant(0) & "' And " & _
                   "ReceiptView = True And " & _
                   "tlbTransactionTypes.TYPE_ID = tlbReceipt.Type " & _
             "Order By TransactionID"
   Set rdoRst1 = rdoConn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
'
'MsgBox rdoRst1!Type
   iRow = 1
   While Not rdoRst1.EOF
      flxSPayment.TextMatrix(iRow, 0) = rdoRst1!TransactionID
      flxSPayment.TextMatrix(iRow, 1) = rdoRst1!description
      flxSPayment.TextMatrix(iRow, 2) = rdoRst1!SageAccountNumber
      flxSPayment.TextMatrix(iRow, 3) = rdoRst1!UnitID
      flxSPayment.TextMatrix(iRow, 4) = Format(rdoRst1!DDate, "dd/mm/yyyy")
      flxSPayment.TextMatrix(iRow, 5) = rdoRst1!Ref
      flxSPayment.TextMatrix(iRow, 6) = rdoRst1!Details
      flxSPayment.TextMatrix(iRow, 7) = Format(rdoRst1!Amount, "0.00")
      flxSPayment.TextMatrix(iRow, 8) = Format(rdoRst1!OSAmount, "0.00")
      flxSPayment.TextMatrix(iRow, 9) = "0.00"
      flxSPayment.TextMatrix(iRow, 10) = IIf(IsNull(rdoRst1!Discount) Or IIf(Not IsNull(rdoRst1!Discount), 0, rdoRst1!Discount) = 0, "0.00", rdoRst1!Discount)
      flxSPayment.TextMatrix(iRow, 11) = rdoRst1!DemandRef
      rdoRst1.MoveNext
      If Not rdoRst1.EOF Then flxSPayment.AddItem ""
      iRow = iRow + 1
   Wend

   rdoRst1.Close
   rdoConn.Close
   
   Set rdoRst1 = Nothing
   Set rdoConn = Nothing
End Sub

Private Sub ConfigurFlxPay(conFlxGrid As Control)
   conFlxGrid.Cols = 8

   If conFlxGrid.Rows = 2 Then
      conFlxGrid.ColWidth(0) = 350        'Serial Number
      conFlxGrid.ColWidth(1) = 1000       'Sage Account Number
      conFlxGrid.ColWidth(2) = 800        'Unit Number
      conFlxGrid.ColWidth(3) = 900        'Due Amount
      conFlxGrid.ColWidth(4) = 900        'Paid Amount
      conFlxGrid.ColWidth(5) = 900        'Part/Full
      conFlxGrid.ColWidth(6) = 1000       'Paid Date
'      conFlxGrid.ColWidth(7) = 800    'NC Amt
'      conFlxGrid.ColWidth(8) = 1400    'NN Amt
'      conFlxGrid.ColWidth(9) = 800    'NC VAT
'      conFlxGrid.ColWidth(10) = 1400     'NN VAT
'      conFlxGrid.ColWidth(11) = 800     'NC Tol
'      conFlxGrid.ColWidth(12) = 1400     'NN Tol
'      conFlxGrid.ColWidth(13) = 1          'This column always at the end for keep BATCHID number
      conFlxGrid.ColWidth(7) = 1      'UniqueRefNumber
   End If
   conFlxGrid.Rows = 2
   conFlxGrid.Clear
'
   conFlxGrid.TextMatrix(0, 0) = "SL"
   conFlxGrid.TextMatrix(0, 1) = "Sage A/C"
   conFlxGrid.TextMatrix(0, 2) = "Unit"
   conFlxGrid.TextMatrix(0, 3) = "Due Amt"
   conFlxGrid.TextMatrix(0, 4) = "Paid Amt"
   conFlxGrid.TextMatrix(0, 5) = "P/F Paid"
   conFlxGrid.TextMatrix(0, 6) = "Paid Dt"
'   conFlxGrid.TextMatrix(0, 7) = "NC Amt"
'   conFlxGrid.TextMatrix(0, 8) = "NN Amt"
'   conFlxGrid.TextMatrix(0, 9) = "NC Vat"
'   conFlxGrid.TextMatrix(0, 10) = "NN Vat"
'   conFlxGrid.TextMatrix(0, 11) = "NC Tol"
'   conFlxGrid.TextMatrix(0, 12) = "NN Tol"
'
   conFlxGrid.RowHeightMin = 315
End Sub

Private Sub txtAddNewAmount_Change()
   If txtAddNewAmount.text = "txtAddNewAmount" Or txtAddNewAmount.text = "" Then Exit Sub
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel + 1) = Format(CCur(txtAddNewAmount.text) * (fVAT_Rate / 100), "0.00")
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel + 2) = Format(CStr(CCur(txtAddNewAmount.text) + CCur(flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, flxAddNewDemands.ColSel + 1))), "0.00")
End Sub

Private Sub txtAddNewAmount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then txtAddNewAmount_LostFocus
   If txtAddNewAmount.text = "" Then Exit Sub
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtAddNewAmount_LostFocus()
   txtAddNewAmount.Visible = False
   If txtAddNewAmount.text = "" Then Exit Sub
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 8) = Format(txtAddNewAmount.text, "0.00")
   CalSubTotal flxAddNewDemands, txtSubTAmount, txtSubTVAT, txtSubTTotal
End Sub

Private Sub txtAddNewDescription_LostFocus()
   txtAddNewDescription.Visible = False
   If txtAddNewDescription.text = "" Then Exit Sub
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Row, 4) = txtAddNewDescription.text

   flxAddNewDemands.ColSel = 5         'From Date
   flxAddNewDemands_Click
End Sub

Private Sub txtBkAc_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdBkList_Click (tabPayment.Tab - 1)
   End If
End Sub

Private Sub txtCCBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdCCBk_Click (tabPayment.Tab - 1)
   End If
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
   If txtDate.text <> "" Then TextBoxFormatDate txtDate
   flxAddNewDemands.TextMatrix(flxAddNewDemands.Rows - 1, flxAddNewDemands.ColSel) = txtDate
   txtDate.Visible = False

   flxAddNewDemands.Col = flxAddNewDemands.ColSel + 1
   flxAddNewDemands_Click
End Sub

Private Sub txtDateBk_GotFocus(Index As Integer)
   If Len(txtDateBk(tabPayment.Tab - 1).text) < 10 Then txtDateBk(tabPayment.Tab - 1).text = Format(Date, "dd/mm/yyyy")
   txtDateBk(tabPayment.Tab - 1).SelStart = 0
   txtDateBk(tabPayment.Tab - 1).SelLength = Len(txtDateBk(tabPayment.Tab - 1).text)
End Sub

Private Sub txtDeptBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdDeptBk_Click (tabPayment.Tab - 1)
   End If
End Sub

Private Sub txtEditAmount_Change()
   If txtEditAmount.text = "txtEditAmount" Or txtEditAmount.text = "" Then Exit Sub
   flxEditDemand.TextMatrix(flxEditDemand.Row, flxEditDemand.ColSel + 1) = _
                  Format(CCur(txtEditAmount.text) * (fVAT_Rate / 100), "0.00")
   flxEditDemand.TextMatrix(flxEditDemand.Row, flxEditDemand.ColSel + 2) = _
                  Format(CStr(CCur(txtEditAmount.text) + CCur(flxEditDemand.TextMatrix(flxEditDemand.Row, _
                  flxEditDemand.ColSel + 1))), "0.00")
End Sub

Private Sub txtEditAmount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then txtEditAmount_LostFocus
   If txtEditAmount.text = "" Then Exit Sub
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtEditAmount_LostFocus()
   txtEditAmount.Visible = False
   If txtEditAmount.text = "" Then Exit Sub
   flxEditDemand.TextMatrix(flxEditDemand.Row, 8) = Format(txtEditAmount.text, "0.00")
   CalSubTotal flxEditDemand, txtEditSubAmount, txtEditSubVat, txtEditSubTotal
End Sub

Private Sub txtEditDescription_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then txtEditDescription_LostFocus
End Sub

Private Sub txtEditDescription_LostFocus()
   txtEditDescription.Visible = False
   If txtEditDescription.text = "" Then Exit Sub
   flxEditDemand.TextMatrix(flxEditDemand.Row, 4) = txtEditDescription.text
End Sub

Private Sub txtEditIssueDate_Click()
   dtEditIssueDate.Visible = True
   dtEditIssueDate.ZOrder 0
   dtEditIssueDate.SetFocus
End Sub

Private Sub txtIssueDate_Change()
   TextBoxChangeDate txtIssueDate
End Sub

Private Sub txtIssueDate_Click()
   If txtTenantName.text = "" Then
      MsgBox "Please choose tenant name.", vbInformation + vbOKOnly, "Tenant Name"
      Exit Sub
   End If
End Sub

Private Sub txtIssueDate_GotFocus()
   If txtTenantName.text = "" Then
      MsgBox "Please choose tenant name.", vbInformation + vbOKOnly, "Tenant Name"
      Exit Sub
   End If
End Sub

Private Sub txtIssueDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtIssueDate, KeyAscii
End Sub

Private Sub txtIssueDate_LostFocus()
   If txtIssueDate.text <> "" Then TextBoxFormatDate txtIssueDate
   flxAddNewDemands.Col = 2
   flxAddNewDemands_Click
'   cboType.SetFocus          i need to change the TextBoxFormatDate, from that method i need to return boolean value to get the status
End Sub

Private Sub txtNCBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdNCBk_Click (tabPayment.Tab - 1)
   End If
End Sub

Private Sub txtNetBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 10 Then txtNetBk_LostFocus (tabPayment.Tab - 1)

   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtNetBk_LostFocus(Index As Integer)
   txtVatBk(tabPayment.Tab - 1).text = Format(IIf(txtNetBk(tabPayment.Tab - 1).text = "", 0, Val(txtNetBk(tabPayment.Tab - 1).text)) * (nTaxCode / 100), "0.00")
   txtNetBk(tabPayment.Tab - 1).text = Format(txtNetBk(tabPayment.Tab - 1).text, "0.00")
End Sub

Private Sub txtNominalCodeTR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdNC_Click
      Exit Sub
   End If
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtNominalCodeTR_LostFocus()
   If txtNominalCodeTR.text = "/" Then txtNominalCodeTR.text = ""
End Sub

Private Sub txtProjBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdProjBk_Click (tabPayment.Tab - 1)
   End If
End Sub

Private Sub txtRecharge_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      lstYNBk(tabPayment.Tab - 1).Top = 840
      lstYNBk(tabPayment.Tab - 1).Left = 9000
      lstYNBk(tabPayment.Tab - 1).Visible = True
      lstYNBk(tabPayment.Tab - 1).SetFocus
      lstYNBk(tabPayment.Tab - 1).ZOrder 0
   End If
End Sub

Private Sub txtRecharge_GotFocus(Index As Integer)
   txtRecharge(tabPayment.Tab - 1).SelStart = 0
   txtRecharge(tabPayment.Tab - 1).SelLength = Len(txtRecharge(tabPayment.Tab - 1).text)
End Sub

Private Sub txtSPayment_Click()
   txtSPayment.SelStart = 0
   txtSPayment.SelLength = Len(txtSPayment.text)
End Sub

Private Sub txtSPayment_GotFocus()
   txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) - CCur(txtSPayment.text), "0.00")
   txtChqNo.text = Format(CCur(txtChqNo.text) - CCur(txtSPayment.text), "0.00")
   txtSPayment.SelStart = 0
   txtSPayment.SelLength = Len(txtSPayment.text)
   iCurRow = flxSPayment.Row
End Sub

Private Sub txtSPayment_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 27 Then
'      cmdSPSave.SetFocus
'      Exit Sub
'   End If

'   If KeyAscii = 13 Or KeyAscii = 10 Then cmdSPSave.SetFocus

   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtSPayment_LostFocus()
   flxSPayment.TextMatrix(iCurRow, 9) = Format(IIf(txtSPayment.text = "", 0, txtSPayment.text), "0.00")
   txtSPaymentTotal.text = Format(CCur(txtSPaymentTotal.text) + CCur(txtSPayment.text), "0.00")
   txtChqNo.text = Format(CCur(txtChqNo.text) + CCur(txtSPayment.text), "0.00")
   If flxSPayment.TextMatrix(iCurRow, 8) > 0 Then baChangesMade(iCurRow) = True
   txtSPayment.text = "0.00"

'   SupTotalPay

   txtSPayment.ZOrder 1
'   cmdSPSave.SetFocus
   bChangesMade = True     'there some changes made in the form
End Sub

Private Sub txtSPDate_GotFocus()
   dtSPDate.Visible = True
   dtSPDate.ZOrder 0
   txtSPDate.SelStart = 0
   txtSPDate.SelLength = Len(txtSPDate.text)
End Sub

Private Sub txtSearchAC_Change()
'   Dim iRow As Integer, iRows As Integer, iTxtLen As Integer
'
'   iTxtLen = Len(txtSearchAC.text)
'   iRows = 1
'   For iRow = 1 To flxPayee.Rows - 1
'      If UCase(txtSearchAC.text) = UCase(Left(flxPayee.TextMatrix(iRow, 0), iTxtLen)) Then
'         flxPayee.RowPosition(iRow) = iRows
'         iRows = iRows + 1
'      End If
'   Next iRow
End Sub

Private Sub txtSPDate_LostFocus()
   dtSPDate.Visible = False
'   flxSPayment.Col = 8
'   flxSPayment_Click
End Sub

Private Sub txtTenantName_KeyPress(KeyAscii As Integer)
   If iSelectedDemandsRow Then Exit Sub
   If KeyAscii = 27 Then cboTenant.Visible = True
End Sub

Private Sub txtTypeBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdTransListBk_Click (tabPayment.Tab - 1)
   End If
End Sub

Private Sub txtUnitBk_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 47 Then
      cmdUnitListBk_Click (tabPayment.Tab - 1)
   End If
End Sub

