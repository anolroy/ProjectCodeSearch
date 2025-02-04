VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmImpTransFromSt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload Bank Statements"
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14475
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpTransFromSt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCBA 
      BackColor       =   &H80000013&
      Caption         =   "Change Bank Account"
      Height          =   375
      Left            =   10785
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   620
      Width           =   1890
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   10395
      TabIndex        =   48
      Top             =   20
      Width           =   10395
      Begin VB.OptionButton optBoth 
         Caption         =   "Both"
         Height          =   255
         Left            =   9720
         TabIndex        =   5
         Top             =   525
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optPayment 
         Caption         =   "Payment"
         Height          =   255
         Left            =   8760
         TabIndex        =   4
         Top             =   525
         Width           =   975
      End
      Begin VB.TextBox txtAccountName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         Height          =   285
         Left            =   7755
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select &Bank Statement"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   525
         Width           =   2490
      End
      Begin VB.TextBox txtBankSt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   525
         Width           =   2970
      End
      Begin VB.OptionButton optReceipt 
         Caption         =   "Receipt"
         Height          =   255
         Left            =   7800
         TabIndex        =   3
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   555
      End
      Begin MSForms.ComboBox cboClientID 
         Height          =   315
         Left            =   600
         TabIndex        =   0
         Top             =   120
         Width           =   2730
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "4815;556"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1763"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   53
         Top             =   120
         Width           =   495
      End
      Begin MSForms.ComboBox cboBC 
         Height          =   315
         Left            =   3720
         TabIndex        =   1
         Top             =   120
         Width           =   2970
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "5239;556"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   5
         ListRows        =   20
         cColumnInfo     =   5
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;10583;0;0;0"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   22
         Left            =   7200
         TabIndex        =   52
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Show:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   7200
         TabIndex        =   51
         Top             =   525
         Width           =   495
      End
   End
   Begin VB.PictureBox picFund 
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
      Left            =   6960
      ScaleHeight     =   3000
      ScaleWidth      =   4185
      TabIndex        =   41
      Top             =   8160
      Visible         =   0   'False
      Width           =   4215
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
         Index           =   0
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   300
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1100
         TabIndex        =   42
         Top             =   300
         Visible         =   0   'False
         Width           =   3045
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFundList 
         Height          =   2370
         Left            =   45
         TabIndex        =   45
         Top             =   600
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   4
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
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   3
         Left            =   1095
         TabIndex        =   47
         Top             =   60
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   75
         Width           =   210
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
         Top             =   60
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdFund 
      BackColor       =   &H00FFC0C0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   40
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   325
   End
   Begin VB.CommandButton cmdRemTrans 
      Caption         =   "&Remove Transaction"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   7365
      Width           =   1935
   End
   Begin VB.CommandButton cmdImportTransactions 
      Caption         =   "&Process Bank Statement"
      Height          =   375
      Left            =   6510
      TabIndex        =   8
      Top             =   7365
      Width           =   1935
   End
   Begin VB.CommandButton cmdSDS 
      Caption         =   "Select a &Different Statement"
      Height          =   375
      Left            =   13560
      TabIndex        =   39
      Top             =   6360
      Width           =   2175
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
      Left            =   1200
      ScaleHeight     =   3000
      ScaleWidth      =   5625
      TabIndex        =   32
      Top             =   8160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtSupplierSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1100
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
         Width           =   4005
      End
      Begin VB.TextBox txtSupplierSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Visible         =   0   'False
         Width           =   900
      End
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
         TabIndex        =   33
         Top             =   15
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplierList 
         Height          =   2370
         Left            =   40
         TabIndex        =   36
         Top             =   600
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   4
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
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   75
         Width           =   165
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   1095
         TabIndex        =   37
         Top             =   60
         Width           =   405
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
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12000
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   30
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
      Begin MSForms.ComboBox cboType 
         Height          =   315
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   1290
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2275;556"
         BoundColumn     =   0
         TextColumn      =   1
         ListRows        =   20
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   14737632
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   12720
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdAccount 
      BackColor       =   &H00FFC0C0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   325
   End
   Begin VB.TextBox txtAcBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Height          =   285
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   13560
      TabIndex        =   13
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnReconTran 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dis&play"
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   11460
      TabIndex        =   11
      Top             =   7365
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   7365
      Width           =   1280
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBank 
      Height          =   5985
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   10557
      _Version        =   393216
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
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   27
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   26
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   25
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Account No."
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   24
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   6
      Left            =   13560
      TabIndex        =   22
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label lblPCB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   13560
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblCB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   13560
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Balance:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   9
      Left            =   13560
      TabIndex        =   19
      Top             =   3360
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Projected Closing Balance:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   8
      Left            =   13560
      TabIndex        =   18
      Top             =   2400
      Width           =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFCFCF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dr"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   6
      Left            =   9535
      TabIndex        =   17
      Top             =   1035
      Width           =   1410
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFCFCF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cr"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   7
      Left            =   10945
      TabIndex        =   16
      Top             =   1035
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   1
      Left            =   1305
      TabIndex        =   15
      Top             =   1075
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   345
      TabIndex        =   14
      Top             =   1075
      Width           =   375
   End
End
Attribute VB_Name = "frmImpTransFromSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szAllBankBalance As String
Private the_array() As String
Private whole_file As String
Private bLoadedSavedTransactions As Boolean
Private bSortingCol() As Boolean, iCurRow  As Integer
Dim szaTenantBalance() As String, iFundRow As Integer

Private Sub EnaDisFormNotFram(szMode As String)
   Dim ctrl As Control

   For Each ctrl In Me
      If ctrl.Name <> "fraAddTran" And _
            ctrl.Container.Name <> "fraAddTran" And _
            TypeName(ctrl) <> "Shape" And _
            TypeName(ctrl) <> "Line" Then
         ctrl.Enabled = IIf(szMode = "Disable", False, True)
      End If
   Next ctrl
End Sub

Private Sub cmdAccount_Click()
   Call PrepareList

   picSupList.Top = IIf(cmdAccount.Top + picSupList.Height < Me.Height, _
                        cmdAccount.Top, Me.Height - picSupList.Height)
   picSupList.Left = Label2(2).Left
   picSupList.Visible = True
   picSupList.ZOrder 0
   cmdAccount.Visible = False
   flxBank.Enabled = False

'   ComponentEnableExpFrame Me, picSupList, DisableAll
End Sub

Private Sub PrepareList()
   ConfigFlxSupplierList

   If InStr(flxBank.TextMatrix(flxBank.row, 2), "Purchase") > 0 Then LoadFlxSupplierList "Supplier"
   If InStr(flxBank.TextMatrix(flxBank.row, 2), "Sales") > 0 Then
      LoadFlxSupplierList "Lessee"
      UpdateBalance
   End If
   If InStr(flxBank.TextMatrix(flxBank.row, 2), "Bank") > 0 Then LoadFlxSupplierList "Bank"
End Sub

Private Sub ConfigFlxFundList()
   Dim szHeader As String

   flxFundList.Cols = 3
   flxFundList.Clear
   szHeader$ = "|<ID|<Name"
   flxFundList.FormatString = szHeader$
   flxFundList.ColWidth(0) = 0          'Solid column
   flxFundList.ColWidth(1) = Label20(3).Left - Label20(2).Left     'Fund No
   flxFundList.ColWidth(2) = flxFundList.Left + flxFundList.Width - Label20(3).Left - 380       'Fund Name

   flxFundList.Rows = 2
   flxFundList.RowHeight(0) = 0
End Sub

Private Sub ConfigFlxSupplierList()
   Dim szHeader As String

   flxSupplierList.Cols = 5
   flxSupplierList.Clear
   szHeader$ = "|<ID|<Name"
   flxSupplierList.FormatString = szHeader$
   flxSupplierList.ColWidth(0) = 0          'Solid column
   flxSupplierList.ColWidth(1) = Label20(1).Left - Label20(0).Left     'Supplier ID
   flxSupplierList.ColWidth(2) = flxSupplierList.Left + flxSupplierList.Width - Label20(1).Left - 380       'Supplier Name
   flxSupplierList.ColWidth(3) = 0          'Balance
   flxSupplierList.ColWidth(4) = 0          'Balance in array --> array index of szaTenantBalance

   flxSupplierList.Rows = 2
   flxSupplierList.RowHeight(0) = 0
End Sub

'  Get all lessees' Account History in an array
Private Sub TenantAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From Tenants;"
   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRptDr.EOF Then
      adoRptDr.Close
      Set adoRptDr = Nothing
      Exit Sub
   End If

   ReDim szaTenantBalance(2, adoRptDr.Fields.Item(0).Value) As String
   adoRptDr.Close

   szSQL = "SELECT Tenants.SageAccountNumber, X.Dr " & _
           "FROM Tenants LEFT OUTER JOIN (" & _
               "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
               "FROM tlbReceipt AS Rpt " & _
               "WHERE Type = 1 OR Type = 23 " & _
               "GROUP BY Rpt.SageAccountNumber) AS X ON " & _
           "Tenants.SageAccountNumber = X.SageAccountNumber;"
'Debug.Print szSQL

   adoRptDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoRptDr.EOF
      szaTenantBalance(0, iIndex) = adoRptDr.Fields.Item("SageAccountNumber").Value
      szaTenantBalance(1, iIndex) = IIf(IsNull(adoRptDr.Fields.Item("Dr").Value), 0, adoRptDr.Fields.Item("Dr").Value)
      szaTenantBalance(2, iIndex) = szaTenantBalance(1, iIndex)
      iIndex = iIndex + 1
      adoRptDr.MoveNext
   Wend

   adoRptDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbReceipt AS Rpt " & _
           "WHERE Type <> 1 AND Type <> 23 " & _
           "GROUP BY SageAccountNumber;"

   adoRptCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRptCr.EOF
      For i = 0 To iIndex - 1
         If szaTenantBalance(0, i) = adoRptCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i < iIndex Then
         szaTenantBalance(1, i) = szaTenantBalance(1, i) - Val(adoRptCr.Fields.Item("Cr").Value)
         szaTenantBalance(2, i) = szaTenantBalance(1, i)
      Else
         iIndex = iIndex + 1
         szaTenantBalance(0, iIndex) = adoRptCr.Fields.Item("Cr").Value
      End If
      adoRptCr.MoveNext
   Wend

   adoRptCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub

Private Sub LoadFlxSupplierList(szAccount As String)
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   If szAccount = "Supplier" Then                                       'Supplier List generating
      szSQL = "SELECT SupplierID, SupplierName " & _
              "FROM Supplier " & _
              "ORDER BY SupplierName;"
   End If
   If szAccount = "Lessee" Then                                         'Lessee List generating
      szSQL = "SELECT T.SageAccountNumber, T.Name " & _
              "FROM Tenants AS T, LeaseDetails AS L, Units AS U, Property AS P " & _
              "WHERE T.SageAccountNumber = L.SageAccountNumber AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = P.PropertyID "
      If cboClientID.Column(0) <> "ALL" Then
         szSQL = szSQL & "AND P.ClientID = '" & cboClientID.Column(0) & "' "
      End If
   End If
   If szAccount = "Bank" Then                                           'Bank List generating
      szSQL = "SELECT CB.NominalCode, CB.Bank_AC_Name " & _
              "FROM tlbClientBanks AS CB "
              
      If cboClientID.Column(0) = "ALL" Then
         szSQL = szSQL + "WHERE CB.CLIENT_ID <> '' "
              
      Else
         szSQL = szSQL + "WHERE CB.CLIENT_ID = '" & cboClientID.Column(0) & "' "
      End If
      szSQL = szSQL & "ORDER BY CB.NominalCode;"
   End If
   
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRst.EOF Then GoTo NoRes
   
   Dim iRow As Integer
   iRow = 1
   
   While Not rstRst.EOF
      flxSupplierList.TextMatrix(iRow, 1) = rstRst.Fields.Item(0).Value
      flxSupplierList.TextMatrix(iRow, 2) = rstRst.Fields.Item(1).Value
'     Column 3 & 4 will be filled in UpdateBalance method
'      flxSupplierList.ColWidth(3) = 0          'Balance
'      flxSupplierList.ColWidth(4) = 0          'Balance in array --> array index of szaTenantBalance
      
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

Private Sub cmdCBA_Click()
   picMain.Enabled = True
   ConfigFlxBank
   txtBankSt.text = ""
End Sub

Private Sub cmdClear_Click()
   Dim i As Integer

   LoadFileInGrid

   For i = 0 To UBound(szaTenantBalance, 2)
      szaTenantBalance(2, i) = szaTenantBalance(1, i)
   Next i

'   If bLoadedSavedTransactions Then
'      If MsgBox("It will clear the saved transactions. Do you want to proceed?", vbQuestion + vbYesNo, "Bank Reconciliation") = vbYes Then
'         ConfigFlxBank
'         cmdBrowse.Enabled = True
'         txtBankSt.text = ""
'         lblPCB.Caption = ""
'         lblCB.Caption = ""
'
'         Dim adoConn As New ADODB.Connection
'
'         adoConn.Open getConnectionString
'         adoConn.Execute "DELETE * FROM tblBankStatement;"
'         adoConn.Close
'         Set adoConn = Nothing
'      End If
'   End If
End Sub

Public Sub LoadDataExternally(Optional adoConn As ADODB.Connection)
   If cboClientID.text = "" Or cboBC.text = "" Or txtAccountName.text = "" Then Exit Sub

   Dim bConn As Boolean

   On Error GoTo SetConnection

   Debug.Print adoConn.Version

   bConn = True
   GoTo ConnectionSet

SetConnection:
   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString
   bConn = False

ConnectionSet:
   CheckSavedRecon adoConn

   txtAcBal.text = Format(BankAccBalance(adoConn, cboBC.Column(0), cboClientID.Column(0)), "0.00")

   If Not bConn Then
      adoConn.Close
      Set adoConn = Nothing
   End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFund_Click()
   picFund.Top = IIf(cmdFund.Top + picFund.Height < Me.Height, _
                        cmdFund.Top, Me.Height - picFund.Height)
   picFund.Left = Label2(3).Left
   picFund.Visible = True
   picFund.ZOrder 0
   cmdFund.Visible = False
   flxBank.Enabled = False
'   ComponentEnableExpFrame Me, picFund, DisableAll
End Sub

Private Sub cmdFund_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      cmdGridUnitClose2_Click 0
   End If
End Sub

Private Sub cmdFund_LostFocus()
   If Not picFund.Visible And flxBank.row <> iFundRow Then
      cmdFund.SetFocus
      ShowMsgInTaskBar "You must select the Fund for the current selection.", "Y", "N"
   End If
End Sub

Private Sub flxFundList_Click()
   If flxFundList.TextMatrix(flxFundList.row, 1) = "" Then Exit Sub

   flxBank.TextMatrix(iFundRow, 4) = flxFundList.TextMatrix(flxFundList.row, 1)

   picFund.Visible = False
   flxBank.Enabled = True
'   ComponentEnableExpFrame Me, picFund, EnableAll
End Sub

Private Sub cmdGridUnitClose2_Click(Index As Integer)
   If Index = 0 Then
      If frmMMain.rtxtMessageDisplay.text = "Part of the receipt will be booked as Receipt on Account." Then
         If MsgBox("Do you wish to cancel Fund selection?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
            flxBank.TextMatrix(iFundRow, 2) = ""
            flxBank.TextMatrix(iFundRow, 3) = ""
            flxBank.TextMatrix(iFundRow, 4) = ""
            flxBank.TextMatrix(iFundRow, 9) = ""
            cmdFund.Visible = False
            flxBank.Enabled = True
            picFund.Visible = False

            szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4)) = _
               Val(szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4))) + _
               IIf(flxBank.TextMatrix(iFundRow, 7) = "", 0, Val(flxBank.TextMatrix(iFundRow, 7)))
         Else
            Exit Sub
         End If
      Else
         picFund.Visible = False
         flxBank.Enabled = True
      End If
   End If

   If Index = 1 Then
      picSupList.Visible = False
      flxBank.Enabled = True
   End If
End Sub

Private Sub cmdImportTransactions_Click()
   Dim i As Integer

   If cmdFund.Visible And cmdFund.Enabled Then
      ShowMsgInTaskBar "You must select the fund to continue.", "Y", "N"
      Exit Sub
   End If

'  ############################### Unblock the following code segment when testing complete
   For i = 1 To flxBank.Rows - 1
      If (flxBank.TextMatrix(i, 2) = "" Or flxBank.TextMatrix(i, 3) = "") And _
            flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) > 0 Then
         ShowMsgInTaskBar "You should select type and account number for all transactions.", "Y", "N"
         Exit Sub
      End If
   Next i
'-------------------------------------------------------------------------------------------------------


'   For i = 1 To flxBank.Rows - 1
'      If flxBank.TextMatrix(i, 2) <> "" And flxBank.TextMatrix(i, 3) <> "" Then
'         If InStr(flxBank.TextMatrix(i, 2), "Bank") > 0 Then
'            If flxBank.TextMatrix(i, 4) <> "" Then HighLightRowsFlxGrid flxBank, i
'         Else
'            HighLightRowsFlxGrid flxBank, i
'         End If
'      End If
'   Next i

   If MsgBox("Marked transactions will be imported into the system." & Chr(13) & _
             "Do you want to continue?", vbQuestion + vbYesNo, _
             "Import Transactions from Bank Statement") = vbNo Then Exit Sub

   ImportTransactions
   cmdCBA_Click
   ShowMsgInTaskBar "Transactions have been uploaded successfully.", "Y", "P"
End Sub

Private Sub ImportTransactions()
   Dim adoConn As New ADODB.Connection
   Dim rstSet As New ADODB.Recordset, rstSplit As New ADODB.Recordset
   Dim adoRst As New ADODB.Recordset, adoRstSpl As New ADODB.Recordset
   Dim adoSplits As New ADODB.Recordset, adoRptTrans As New ADODB.Recordset
   Dim i As Integer, szSQL As String, lSlNumber As Long
   Dim lRpt_ID As Long, bFlag As Boolean, iSIdx As Integer 'Split Index
   Dim cTotalAllocAmt As Currency, cAllocAmt As Currency

   adoConn.Open getConnectionString

   With rstSet
      bFlag = False

      For i = 1 To flxBank.Rows - 1
'-------------------------------------------------------------------------------------------------------------
'***************************           BANK RECEIPTS           ***********************************************
'-------------------------------------------------------------------------------------------------------------
         If flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) > 0 Then
            If flxBank.TextMatrix(i, 2) = "Bank Receipts" Then                   'Bank Receipts
               If Not bFlag Then
                  szSQL = "SELECT * FROM tlbBankPayment"
                  .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                  bFlag = True
               End If
               .AddNew
               !MY_ID = Format(Now, "yyyymmddhhmmss") & CStr(i)
               !TRAN_ID = SlNumber("BR", "tlbBankPayment", adoConn)
               !BANK_AC = flxBank.TextMatrix(i, 3)
               !TRANS = "BR"
               !TRAN_DATE = Format(flxBank.TextMatrix(i, 1), "DD MMMM YYYY")
               !clientID = cboClientID.Column(0)
               !NOMINAL_CODE = flxBank.TextMatrix(i, 3)
               !PROJ_REF = flxBank.TextMatrix(i, 6)              'Reference
               !DEPT_ID = flxBank.TextMatrix(i, 4)               'Fund
               !description = flxBank.TextMatrix(i, 5)
               !NET_AMOUNT = CCur(flxBank.TextMatrix(i, 8))
               !VAT = 0
               !TAX_CODE = "T9"
               !TransactionType = 12     'Bank Receipt = sdoBR 12

'issue 523
'Modified by anol 2015
               UpdateBankAcBal_Plus adoConn, !NET_AMOUNT, flxBank.TextMatrix(i, 3), cboClientID.Value
               .Update
            End If
'-------------------------------------------------------------------------------------------------------------
'***************************           BANK PAYMENTS           ***********************************************
'-------------------------------------------------------------------------------------------------------------
            If InStr(flxBank.TextMatrix(i, 2), "Bank Payments") > 0 Then      'Bank Payments
               If Not bFlag Then
                  szSQL = "SELECT * FROM tlbBankPayment"
                  .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                  bFlag = True
               End If
               .AddNew
               !MY_ID = Format(Now, "yyyymmddhhmmss") & CStr(i)
               !TRAN_ID = SlNumber("BR", "tlbBankPayment", adoConn)
               !BANK_AC = flxBank.TextMatrix(i, 3)
               !TRANS = "BP"
               !TRAN_DATE = Format(flxBank.TextMatrix(i, 1), "DD MMMM YYYY")
               !clientID = cboClientID.Column(0)
               !NOMINAL_CODE = flxBank.TextMatrix(i, 3)
               !PROJ_REF = flxBank.TextMatrix(i, 6)              'Reference
               !DEPT_ID = flxBank.TextMatrix(i, 4)               'Fund
               !description = flxBank.TextMatrix(i, 5)
               !NET_AMOUNT = CCur(flxBank.TextMatrix(i, 7))
               !VAT = 0
               !TAX_CODE = "T9"
               !TransactionType = 11     'Bank Payment = sdoBP 11
'issue 523
'Modified by anol 26 Jan 2015
               UpdateBankAcBal_Minus adoConn, !NET_AMOUNT, flxBank.TextMatrix(i, 3), cboClientID.Value

               .Update
            End If
         End If
      Next i
      If bFlag Then
         .Close
         bFlag = False
      End If
'-------------------------------------------------------------------------------------------------------------
'***************************           SALES REFUND            ***********************************************
'-------------------------------------------------------------------------------------------------------------
      For i = 1 To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) > 0 Then
            If InStr(flxBank.TextMatrix(i, 2), "Sales Refunds") > 0 Then                  'Sales Refunds
               lRpt_ID = SlNumber("TI", "tlbReceipt", adoConn)
               Exit For
            End If
         End If
      Next i

      For i = i To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) > 0 Then
            If flxBank.TextMatrix(i, 2) = "Sales Refunds" Then                              'Sales Refunds
               If Not bFlag Then
                  szSQL = "SELECT * FROM tlbReceipt;"
                  .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                  bFlag = True

'    Saving the split(s) of the header
                  szSQL = "SELECT * FROM tlbReceiptSplit;"
                  rstSplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
               End If

               .AddNew
               !TransactionID = lRpt_ID
               !Type = 23
               !SageAccountNumber = flxBank.TextMatrix(i, 3)
               !unitid = GetUnitIDbyTenantID(flxBank.TextMatrix(i, 3), adoConn)
               !RDate = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               !dDate = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               !Ref = "SA" & Format(Now, "yymmddhhmmss")
               !Details = "Sales Receipt Refund"
               !Amount = flxBank.TextMatrix(i, 7)
               !OSAmount = !Amount
               !ReceiptView = True
               !BankCode = txtAccountName.text
               !nominalCode = !BankCode
               !ExtRef = flxBank.TextMatrix(i, 6)
               !SlNumber = SlNumber("SRR", "tlbReceipt", adoConn)
               !fundID = flxBank.TextMatrix(i, 4)
               .Update

'    Saving the split(s) of the header
               rstSplit.AddNew
               rstSplit.Fields.Item("TransactionID").Value = UniqueID()
               rstSplit.Fields.Item("RptHeader").Value = lRpt_ID
               rstSplit.Fields.Item("FundID").Value = flxBank.TextMatrix(i, 4)
               rstSplit.Fields.Item("Amount").Value = flxBank.TextMatrix(i, 7)
               rstSplit.Fields.Item("SplitID").Value = 1
               rstSplit.Fields.Item("DueDate").Value = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               rstSplit.Fields.Item("Description").Value = "Sales Receipt Refund"
               rstSplit.Update

               lRpt_ID = lRpt_ID + 1
            End If
         End If
      Next i
      If bFlag Then
         .Close
         rstSplit.Close
         bFlag = False
      End If
'-------------------------------------------------------------------------------------------------------------
'***************************           PURCHASE REFUND            ********************************************
'-------------------------------------------------------------------------------------------------------------
      For i = 1 To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) > 0 Then
            If InStr(flxBank.TextMatrix(i, 2), "Purchase Refunds") > 0 Then      'Purchase Refunds
               lRpt_ID = SlNumber("TI", "tlbPayment", adoConn)
               Exit For
            End If
         End If
      Next i

      For i = i To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) > 0 Then
            If InStr(flxBank.TextMatrix(i, 2), "Purchase Refunds") > 0 Then      'Purchase Refunds
               If Not bFlag Then
                  szSQL = "SELECT * FROM tlbPayment;"
                  .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                  bFlag = True
               End If

               .AddNew
               !TransactionID = lRpt_ID
               lRpt_ID = lRpt_ID + 1

               !Type = 24
               !SageAccountNumber = flxBank.TextMatrix(i, 3)
               !PDate = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               !Ref = "PPR" & Format(Now, "yymmddhhmmss")
               !Details = "Purchase Payment Refund"
               !Amount = flxBank.TextMatrix(i, 8)
               !OSAmount = !Amount
               !PaymentView = True
               !BankCode = txtAccountName.text
               !nominalCode = !BankCode
               !ExtRef = flxBank.TextMatrix(i, 6)
               !SlNumber = SlNumber("PPR", "tlbPayment", adoConn)
               !fundID = flxBank.TextMatrix(i, 4)

               .Update
            End If
         End If
      Next i
      If bFlag Then
         .Close
         rstSplit.Close
         bFlag = False
      End If
'-------------------------------------------------------------------------------------------------------------
'***************************           SALES RECEIPTS             ********************************************
'-------------------------------------------------------------------------------------------------------------
      For i = 1 To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) > 0 Then
            '
            'Simple Receipt --> where the receipt amount <= Total Invoice amount.
            'The Receipt will be fully allocated against invoice(s)
            If flxBank.TextMatrix(i, 2) = "Sales Receipts" And Val(flxBank.TextMatrix(i, 10)) <= 0 Then
               szSQL = "SELECT * FROM tlbReceipt;"
               .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

               'Create Receipt Header
               lSlNumber = SlNumber("SR", "tlbReceipt", adoConn)
               .AddNew
               lRpt_ID = SlNumber("TI", "tlbReceipt", adoConn)
               .Fields.Item("TransactionID").Value = lRpt_ID
               .Fields.Item("Type").Value = 3
               .Fields.Item("SageAccountNumber").Value = flxBank.TextMatrix(i, 3)
               .Fields.Item("UnitID").Value = GetUnitIDbyTenantID(flxBank.TextMatrix(i, 3), adoConn)
               .Fields.Item("RDate").Value = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               .Fields.Item("DDate").Value = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               .Fields.Item("Ref").Value = "SR" & Format(Now, "yymmddhhmmss")
               .Fields.Item("Details").Value = "Receipt"
               .Fields.Item("Amount").Value = Val(flxBank.TextMatrix(i, 7))
               .Fields.Item("OSAmount").Value = 0
               .Fields.Item("ReceiptView").Value = False
               .Fields.Item("BankCode").Value = txtAccountName.text
               .Fields.Item("NominalCode").Value = .Fields.Item("BankCode").Value
               .Fields.Item("ExtRef").Value = flxBank.TextMatrix(i, 6)
               .Fields.Item("SlNumber").Value = lSlNumber
               .Fields.Item("FundID").Value = 0
               .Fields.Item("RptAmtType").Value = "CHQ"
               .Update
               .Close

               cAllocAmt = Val(flxBank.TextMatrix(i, 7))

               'Collect all Invoices of the specific lessee
               szSQL = "SELECT * FROM tlbReceipt " & _
                       "WHERE (Type = 1 OR Type = 23) AND OSAmount > 0 AND " & _
                           "SageAccountNumber = '" & flxBank.TextMatrix(i, 3) & "' " & _
                       "ORDER BY RDate;"
'Debug.Print szSQL
               adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

               '
               'Create Receipt Splits     **********
               '
               iSIdx = 1
               szSQL = "SELECT * FROM tlbReceiptSplit;"
               adoSplits.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

               While Not adoRst.EOF And cAllocAmt > 0
                  '1st: Get all splits of the SI
                  szSQL = "SELECT S.* " & _
                          "FROM tlbReceiptSplit AS S " & _
                          "WHERE S.RptHeader = " & adoRst.Fields.Item("TransactionID").Value & " AND " & _
                              "S.OSAmount > 0;"
'Debug.Print szSQL
                  adoRstSpl.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

                  cTotalAllocAmt = 0
                  While Not adoRstSpl.EOF And cAllocAmt > 0
                     adoSplits.AddNew
                     adoSplits.Fields.Item("TransactionID").Value = UniqueID()
                     adoSplits.Fields.Item("RptHeader").Value = lRpt_ID
                     adoSplits.Fields.Item("FundID").Value = adoRstSpl.Fields.Item("FundID").Value
                     If Val(adoRstSpl.Fields.Item("OSAmount").Value) <= cAllocAmt Then
                        adoSplits.Fields.Item("Amount").Value = adoRstSpl.Fields.Item("OSAmount").Value
                        cAllocAmt = cAllocAmt - adoSplits.Fields.Item("Amount").Value
                     Else
                        adoSplits.Fields.Item("Amount").Value = cAllocAmt
                        cAllocAmt = 0
                     End If
                     cTotalAllocAmt = cTotalAllocAmt + CCur(adoSplits.Fields.Item("Amount").Value)
                     adoSplits.Fields.Item("SplitID").Value = iSIdx
                     iSIdx = iSIdx + 1
                     adoSplits.Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
                     adoSplits.Fields.Item("Description").Value = flxBank.TextMatrix(i, 5)
                     adoSplits.Fields.Item("AllocTranID").Value = adoRstSpl.Fields.Item("TransactionID").Value

                     szSQL = "UPDATE tlbReceiptSplit " & _
                             "SET OSAmount = OSAmount - " & CCur(adoSplits.Fields.Item("Amount").Value) & " " & _
                             "WHERE TransactionID = '" & adoSplits.Fields.Item("AllocTranID").Value & "';"
                     adoConn.Execute szSQL

                     szSQL = "UPDATE tlbReceipt " & _
                             "SET OSAmount = OSAmount - " & CCur(adoSplits.Fields.Item("Amount").Value) & ", " & _
                                 "ReceiptView = " & IIf(adoRst.Fields.Item("OSAmount").Value - _
                                                   CCur(adoSplits.Fields.Item("Amount").Value) > 0, True, False) & " " & _
                             "WHERE  TransactionID = " & adoRst.Fields.Item("TransactionID").Value & ";"
                     adoConn.Execute szSQL

                     adoSplits.Update
                     adoRstSpl.MoveNext
                  Wend                    'Not adoRstSpl.EOF And cAllocAmt > 0

                  ' Create the relationship between the SI and SR header.
                  ' Relationship btn SI split and SR split is recorded in receipt split table.
                  szSQL = "SELECT * FROM RptTransactions;"
                  adoRptTrans.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                  adoRptTrans.AddNew
                  adoRptTrans!TranType = "AL"
                  adoRptTrans!TransactionID = SlNumber("TI", "RptTransactions", adoConn)
                  adoRptTrans!Alloc_Unalloc = 1

                  adoRptTrans!FromTran = lRpt_ID      'Receipt ID
                  adoRptTrans!toTran = adoRst.Fields.Item("TransactionID").Value     'It is Invoice ID
                  adoRptTrans!AllocDate = Format(Date, "DD MMMM YYYY")
                  adoRptTrans!ReceiptAmount = cTotalAllocAmt
                  adoRptTrans!BankCode = txtAccountName.text
                  adoRptTrans!nominalCode = adoRptTrans!BankCode
                  adoRptTrans!SlNumber = lSlNumber
                  adoRptTrans.Update
                  adoRptTrans.Close

                  adoRstSpl.Close
                  adoRst.MoveNext
               Wend           'Not adoRst.EOF And cAllocAmt > 0
               adoRst.Close
               adoSplits.Close
            End If

            '
            'Simple RoA --> where the receipt amount [flxBank.TextMatrix(i, 7) = flxBank.TextMatrix(i, 10)]
            'The Receipt is unallocated will be booked as RoA
            If flxBank.TextMatrix(i, 2) = "Sales Receipts" And _
                  Val(flxBank.TextMatrix(i, 10)) = Val(flxBank.TextMatrix(i, 7)) Then
               szSQL = "SELECT * FROM tlbReceipt;"
               .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

               'Create Receipt Header
               lSlNumber = SlNumber("SA", "tlbReceipt", adoConn)
               .AddNew
               lRpt_ID = SlNumber("TI", "tlbReceipt", adoConn)
               .Fields.Item("TransactionID").Value = lRpt_ID
               .Fields.Item("Type").Value = 4
               .Fields.Item("SageAccountNumber").Value = flxBank.TextMatrix(i, 3)
               .Fields.Item("UnitID").Value = GetUnitIDbyTenantID(flxBank.TextMatrix(i, 3), adoConn)
               .Fields.Item("RDate").Value = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               .Fields.Item("DDate").Value = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
               .Fields.Item("Ref").Value = "SA" & Format(Now, "yymmddhhmmss")
               .Fields.Item("Details").Value = "Receipt on Account"
               .Fields.Item("Amount").Value = CCur(flxBank.TextMatrix(i, 7))
               .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
               .Fields.Item("ReceiptView").Value = True
               .Fields.Item("BankCode").Value = txtAccountName.text
               .Fields.Item("NominalCode").Value = .Fields.Item("BankCode").Value
               .Fields.Item("ExtRef").Value = flxBank.TextMatrix(i, 6)
               .Fields.Item("SlNumber").Value = lSlNumber
               .Fields.Item("FundID").Value = 0
               .Fields.Item("RptAmtType").Value = "CHQ"
               .Update
               .Close

               '
               'Create Receipt Splits     **********
               '
               szSQL = "SELECT * FROM tlbReceiptSplit;"
               adoSplits.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

               adoSplits.AddNew
               adoSplits.Fields.Item("TransactionID").Value = UniqueID()
               adoSplits.Fields.Item("RptHeader").Value = lRpt_ID
               adoSplits.Fields.Item("FundID").Value = flxBank.TextMatrix(i, 4)
               adoSplits.Fields.Item("Amount").Value = CCur(flxBank.TextMatrix(i, 7))
               adoSplits.Fields.Item("SplitID").Value = 1
               adoSplits.Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
               adoSplits.Fields.Item("Description").Value = flxBank.TextMatrix(i, 5)

               adoSplits.Update
            End If
         End If
      Next i

   End With

   Set rstSet = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub
'
'Private Sub UpdateSIBalance(lRptID As Long, szSpID As String, cAmt As Currency, adoConn As ADODB.Connection)
'   Dim szSQL As String
'
'   szSQL = "UPDATE tlbReceiptSplit " & _
'           "SET OSAmount = OSAmount - " & cAmt & " " & _
'           "WHERE TransactionID = '" & szSpID & "';"
'   adoConn.Execute szSQL
'
'   szSQL = "UPDATE tlbReceipt " & _
'           "SET OSAmount = OSAmount - " & 3000 & ", " & _
'               "ReceiptView = IIf(OSAmount > 0, True, False) " & _
'           "WHERE  TransactionID = " & lRptID & ";"
'   adoConn.Execute szSQL
'End Sub

Private Sub cmdRemTrans_Click()
   Dim i As Integer

   For i = 1 To flxBank.Rows - 1
      If flxBank.TextMatrix(i, 0) = "X" Then
         If flxBank.TextMatrix(i, 2) = "" Then
            flxBank.RowHeight(i) = 0
            flxBank.TextMatrix(i, 0) = "-"
            If picType.Visible Then picType.Visible = False
            If cmdAccount.Visible Then cmdAccount.Visible = False
            If picSupList.Visible Then picSupList.Visible = False
            If picFund.Visible Then picFund.Visible = False

            ShowMsgInTaskBar "Transaction has been removed.", "Y", "P"
            Exit Sub
         Else
            ShowMsgInTaskBar "The transaction will not be removed.", "Y", "N"
            Exit Sub
         End If
      End If
   Next i

   If i = flxBank.Rows Then
      ShowMsgInTaskBar "No Transaction has been selected.", "Y", "N"
   End If
End Sub

Private Sub cmdUnReconTran_Click()
'   If cboClientID.text = "" Then
'      cboClientID.SetFocus
'      Exit Sub
'   End If
'   If cboBC.text = "" Then
'      cboBC.SetFocus
'      Exit Sub
'   End If
'   If txtAccountName.text = "" Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'
''   connect to database
'   adoConn.Open getConnectionString
'
'   CheckSavedRecon adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   cboClientID.Locked = True
'   cboBC.Locked = True
End Sub

Private Sub flxBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxBank.ToolTipText = flxBank.TextMatrix(flxBank.MouseRow, flxBank.MouseCol)
End Sub

Private Sub UpdateBalance()
   Dim i As Integer, j As Integer

'*  flxSupplierList.ColWidth(3) = 0          'Balance
'*  flxSupplierList.ColWidth(4) = 0          'Balance in array --> array index of szaTenantBalance
   With flxSupplierList
      For i = 1 To .Rows - 1
         For j = 0 To UBound(szaTenantBalance, 2) - 1
            If .TextMatrix(i, 1) = szaTenantBalance(0, j) Then
               .TextMatrix(i, 3) = Format(szaTenantBalance(2, j), "0.00")
               .TextMatrix(i, 4) = j
               Exit For
            End If
         Next j
         If j = UBound(szaTenantBalance, 2) Then .TextMatrix(i, 3) = ""
      Next i
   End With
End Sub

Private Sub flxSupplierList_Click()
   If flxSupplierList.TextMatrix(flxSupplierList.row, 1) = "" Then Exit Sub

   If flxBank.TextMatrix(flxBank.row, 2) = "Sales Receipts" Then
      If flxBank.TextMatrix(flxBank.row, 3) <> "" And _
            flxBank.TextMatrix(flxBank.row, 3) <> flxSupplierList.TextMatrix(flxSupplierList.row, 1) Then
'        if user change the lessee's account after they select a lessee
         szaTenantBalance(2, flxSupplierList.TextMatrix(flxBank.TextMatrix(flxBank.row, 9), 4)) = _
            Val(szaTenantBalance(2, flxSupplierList.TextMatrix(flxBank.TextMatrix(flxBank.row, 9), 4))) + _
            IIf(flxBank.TextMatrix(flxBank.row, 7) = "", 0, Val(flxBank.TextMatrix(flxBank.row, 7)))
      End If

      flxBank.TextMatrix(flxBank.row, 3) = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
      flxBank.TextMatrix(flxBank.row, 9) = flxSupplierList.row

      If Val(szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4))) < _
            IIf(flxBank.TextMatrix(flxBank.row, 7) = "", 0, Val(flxBank.TextMatrix(flxBank.row, 7))) Then
         If Val(szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4))) < 0 Then
            flxBank.TextMatrix(flxBank.row, 10) = IIf(flxBank.TextMatrix(flxBank.row, 7) = "", 0, flxBank.TextMatrix(flxBank.row, 7))
         Else
            flxBank.TextMatrix(flxBank.row, 10) = IIf(flxBank.TextMatrix(flxBank.row, 7) = "", 0, flxBank.TextMatrix(flxBank.row, 7)) - _
                                                  szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4))
         End If
      Else
         flxBank.TextMatrix(flxBank.row, 10) = IIf(flxBank.TextMatrix(flxBank.row, 7) = "", 0, flxBank.TextMatrix(flxBank.row, 7)) - _
                                               szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4))
      End If

      szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4)) = _
         IIf(szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4)) = "", _
            0, Val(szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4)))) - _
         IIf(flxBank.TextMatrix(flxBank.row, 7) = "", 0, Val(flxBank.TextMatrix(flxBank.row, 7)))

      If Val(szaTenantBalance(2, flxSupplierList.TextMatrix(flxSupplierList.row, 4))) < 0 Then
         cmdFund.Enabled = True
         ShowMsgInTaskBar "Part of the receipt will be booked as Receipt on Account.", "Y", "P"

         flxBank.col = 4
         flxBank_Click
      Else
         cmdFund.Enabled = False
      End If
   End If
   If flxBank.TextMatrix(flxBank.row, 2) = "Purchase Payments" Then
      flxBank.TextMatrix(flxBank.row, 3) = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
      flxBank.TextMatrix(flxBank.row, 9) = flxSupplierList.row
      cmdFund.Enabled = False
   End If
   If InStr(flxBank.TextMatrix(flxBank.row, 2), "Bank") > 0 Or _
         InStr(flxBank.TextMatrix(flxBank.row, 2), "Refunds") > 0 Then        'Bank Receipts & Bank Payments
                                                                              'Sales Refunds & Purchase Refunds
      flxBank.TextMatrix(flxBank.row, 3) = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
      flxBank.TextMatrix(flxBank.row, 9) = flxSupplierList.row
      cmdFund.Enabled = True
      flxBank.col = 4
      flxBank_Click
   End If

   cmdGridUnitClose2_Click 1
End Sub

Private Sub flxSupplierList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxSupplierList.ToolTipText = flxSupplierList.TextMatrix(flxSupplierList.MouseRow, 3)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub Label2_Click(Index As Integer)
   If Index = 0 Then SortingGrid flxBank, Index + 1, bSortingCol(Index), "Date"
   If Index = 5 Then SortingGrid flxBank, Index + 1, bSortingCol(Index), "Text"
   If Index = 6 Or Index = 7 Then SortingGrid flxBank, Index + 1, bSortingCol(Index), "Currency"

   If Index = 0 Or (Index > 4 And Index < 8) Then
      bSortingCol(Index) = IIf(bSortingCol(Index), False, True)

      LblSortingClicked Index, Label2, 0, 7
   End If
End Sub

Private Sub optBoth_Click()
   ReceiptPayment
End Sub

Private Sub optPayment_Click()
   ReceiptPayment
End Sub

Private Sub optReceipt_Click()
   ReceiptPayment
End Sub

Private Sub txtBankSt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Right(txtBankSt.text, 4) = ".csv" Then LoadFileInGrid
   End If
End Sub

Private Sub cboBC_Click()
   If cboBC.text <> "" Then
      txtAccountName.text = cboBC.Column(0)

      Dim adoConn As New ADODB.Connection

      On Error GoTo ErrorHandler

      adoConn.Open getConnectionString

      txtAcBal.text = BankAccBalance(adoConn, cboBC.Column(0), cboClientID.Column(0))
      txtAcBal.text = Format(txtAcBal.text, "0.00")

      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

ErrorHandler:
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cboBC_GotFocus()
   If cboClientID.ListIndex < 0 Then
      ShowMsgInTaskBar "Please select a client first.", , "N"
      cboClientID.SetFocus
      Exit Sub
   End If
End Sub

Private Sub cboClientID_Click()
'   If cboBC.ListIndex >= 0 Then Exit Sub
   If cboClientID.ListIndex < 0 Then Exit Sub

   Dim adoConn As New ADODB.Connection

   On Error GoTo ErrorHandler

   adoConn.Open getConnectionString

   szAllBankBalance = BankAndBalance(adoConn)

NoRes:
   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdSave_Click()
'   If flxBank.TextMatrix(1, 1) = "" Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim i As Integer
'
''   connect to database
'   adoConn.Open getConnectionString
'
'   adoConn.Execute "DELETE * FROM tblBankStatement;"
'
'   szSQL = "SELECT * FROM tblBankStatement;"
'   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
''   szHeader$ = "X|<Date|<Reference|<Matched Ref.|>Dr|>Cr|>6|7"
'   With adoRst
'      For i = 1 To flxBank.Rows - 1
'         If flxBank.TextMatrix(i, 1) = "" Then Exit For
'
'         .AddNew
'         .Fields.Item("TranDate").Value = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
'         .Fields.Item("ClientBankID").Value = cboBC.Column(2)
'         .Fields.Item("StatementReference").Value = flxBank.TextMatrix(i, 2)
'         .Fields.Item("MatchedRef").Value = IIf(flxBank.TextMatrix(i, 3) = "", "NM", flxBank.TextMatrix(i, 3)) 'NM -> Not Matched
'         .Fields.Item("Dr").Value = IIf(flxBank.TextMatrix(i, 4) = "", 0, flxBank.TextMatrix(i, 4))
'         .Fields.Item("Cr").Value = IIf(flxBank.TextMatrix(i, 5) = "", 0, flxBank.TextMatrix(i, 5))
'         .Update
'      Next i
'      .Close
'   End With
'
'   Set adoRst = Nothing
'   adoConn.Close
'   Set adoConn = Nothing
'
'   ShowMsgInTaskBar "Bank Reconciliation has been saved."
End Sub

Private Function BankAndBalance(adoConn As ADODB.Connection) As String
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   If cboClientID.Column(0) = "ALL" Then
      szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.ClosingBal AS BAL, CB.CLIENT_ID, " & _
                  "CB.PCB " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE CB.NominalCode = N.Code AND CB.CLIENT_ID <> '' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.ClosingBal, " & _
                  "CB.CLIENT_ID, CB.PCB;"
   Else
      szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.ClosingBal AS BAL, CB.CLIENT_ID, " & _
                  "CB.PCB " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE CB.NominalCode = N.Code AND " & _
                  "CB.CLIENT_ID = '" & cboClientID.Column(0) & "' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.ClosingBal, " & _
                  "CB.CLIENT_ID, CB.PCB;"
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "Please setup bank account for the client."
   Else
      ReDim szaData(5, adoRst.RecordCount - 1) As String

      While Not adoRst.EOF
         szaData(0, iRec) = adoRst.Fields.Item("BNC").Value
         szaData(1, iRec) = adoRst.Fields.Item("BNN").Value
         szaData(2, iRec) = adoRst.Fields.Item("ID").Value
         szaData(3, iRec) = adoRst.Fields.Item("CLIENT_ID").Value
         szaData(4, iRec) = IIf(IsNull(adoRst.Fields.Item("BAL").Value), "", adoRst.Fields.Item("BAL").Value)
         szaData(5, iRec) = IIf(IsNull(adoRst.Fields.Item("PCB").Value), "", adoRst.Fields.Item("PCB").Value)
         iRec = iRec + 1
         adoRst.MoveNext
      Wend
      cboBC.Clear
      cboBC.Column() = szaData()
      txtAccountName.text = ""
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Function

Private Sub cmdBrowse_Click()
   If cboClientID.text = "" Or cboBC.text = "" Then
      ShowMsgInTaskBar "Please client and Bank account.", "Y", "N"
      Exit Sub
   End If

   txtBankSt.text = SelectBankStatement

   If txtBankSt.text = "" Then Exit Sub

   Dim file_name As String
   Dim fnum As Integer
    ' Load the file.
   fnum = FreeFile

   Open txtBankSt.text For Input As fnum
   whole_file = Input$(LOF(fnum), #fnum)
   Close fnum

   picMain.Enabled = False
   LoadFileInGrid
   ReceiptPayment
End Sub

Private Sub ReceiptPayment()
   Dim i As Integer

   For i = 1 To flxBank.Rows - 1
      If flxBank.TextMatrix(i, 0) <> "-" And flxBank.RowHeight(i) = 0 Then
         flxBank.RowHeight(i) = 240
      End If
   Next i

   If optReceipt.Value Then
      For i = 1 To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 8) <> "" Then
            flxBank.RowHeight(i) = 0
         End If
      Next i
   End If

   If optPayment.Value Then
      For i = 1 To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 7) <> "" Then
            flxBank.RowHeight(i) = 0
         End If
      Next i
   End If
End Sub

Private Function StGridTotal() As Currency
'   Bank = "X|<Date|<Reference|<Matched Ref.|>Dr|>Cr|>6|7"
   Dim i As Integer

   StGridTotal = 0
   For i = 1 To flxBank.Rows - 1
      If flxBank.TextMatrix(i, 5) <> "" Then
         StGridTotal = StGridTotal + CCur(flxBank.TextMatrix(i, 5))
      Else
         StGridTotal = StGridTotal - CCur(flxBank.TextMatrix(i, 4))
      End If
   Next i
End Function

Private Sub LoadFileInGrid()
   Dim lines As Variant
   Dim one_line As Variant
   Dim num_rows As Long
   Dim num_cols As Long
   Dim r As Long
   Dim c As Long

   ' Break the file into lines.
   lines = Split(whole_file, vbCrLf)

   ' Dimension the array.
   num_rows = UBound(lines)
   one_line = Split(lines(0), ",")
   num_cols = UBound(one_line)
   ReDim the_array(num_rows, num_cols)

   ' Copy the data into the array.
   For r = 0 To num_rows
      If Len(lines(r)) > 0 Then
         one_line = Split(lines(r), ",")
         For c = 0 To num_cols
            the_array(r, c) = one_line(c)
         Next c
      End If
   Next r

   flxBank.Rows = 2
   flxBank.Clear

'  Transfer the data from array to grid
   For r = 1 To num_rows
      flxBank.TextMatrix(r, 1) = the_array(r - 1, 0)
      flxBank.TextMatrix(r, 6) = the_array(r - 1, 1)
      If Val(the_array(r - 1, 2)) < 0 Then
         flxBank.TextMatrix(r, 8) = Format(the_array(r - 1, 2) * (-1), "0.00")
      Else
         flxBank.TextMatrix(r, 7) = Format(the_array(r - 1, 2), "0.00")
      End If
      If r < num_rows Then flxBank.AddItem ""
   Next r

   flxBank.col = 1
   flxBank.Sort = 1
   ShowMsgInTaskBar "Bank statement has been uploaded in the grid successfully"
End Sub

Private Sub CheckSavedRecon(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim r As Integer, i As Integer

   szSQL = "SELECT * FROM tblBankStatement WHERE ClientBankID = " & cboBC.Column(2) & ";"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      bLoadedSavedTransactions = True
      cmdBrowse.Enabled = False
      flxBank.Clear
      flxBank.Rows = 2
      r = 1
      
      While Not adoRst.EOF
         flxBank.TextMatrix(r, 1) = Format(adoRst.Fields.Item("TranDate").Value, "DD/MM/YYYY")
         flxBank.TextMatrix(r, 2) = adoRst.Fields.Item("StatementReference").Value
         flxBank.TextMatrix(r, 3) = IIf(adoRst.Fields.Item("MatchedRef").Value = "NM", "", adoRst.Fields.Item("MatchedRef").Value)
         flxBank.TextMatrix(r, 4) = IIf(adoRst.Fields.Item("Dr").Value = 0, "", _
                                        Format(adoRst.Fields.Item("Dr").Value, "0.00"))
         flxBank.TextMatrix(r, 5) = IIf(adoRst.Fields.Item("Cr").Value = 0, "", _
                                        Format(adoRst.Fields.Item("Cr").Value, "0.00"))
         If flxBank.TextMatrix(r, 3) <> "" Then
            If flxBank.TextMatrix(r, 4) <> "" Then _
               lblCB.Caption = Format(CCur(lblCB.Caption) - CCur(flxBank.TextMatrix(r, 4)), "0.00")
            If flxBank.TextMatrix(r, 5) <> "" Then _
               lblCB.Caption = Format(CCur(lblCB.Caption) + CCur(flxBank.TextMatrix(r, 5)), "0.00")
            HighLightRowsFlxGrid flxBank, r
         End If
         r = r + 1
         adoRst.MoveNext
         If Not adoRst.EOF Then flxBank.AddItem ""
      Wend
   Else
      cmdBrowse.Enabled = True
      bLoadedSavedTransactions = False
   End If
   
   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub ConfigFlxBank()
   Dim i As Integer
   Dim szHeader As String

   flxBank.Clear
   szHeader$ = "X|<Date|<Type|<AccountNo|<Fund|<Description|<Reference|>Dr|>Cr|9|10"
   flxBank.FormatString = szHeader$

   flxBank.Rows = 2
   flxBank.Cols = 11
   flxBank.RowHeight(0) = 0

   flxBank.ColWidth(0) = Label2(0).Left - flxBank.Left

   For i = 1 To flxBank.Cols - 4
      flxBank.ColWidth(i) = Label2(i).Left - Label2(i - 1).Left
   Next i
   flxBank.ColWidth(i) = flxBank.Left + flxBank.Width - Label2(i - 1).Left - 380
   flxBank.ColWidth(flxBank.Cols - 2) = 0             '9
   flxBank.ColWidth(flxBank.Cols - 1) = 0             '10
End Sub

Private Function UnclearedBalance() As Currency
   Dim iRow As Integer
   On Error Resume Next

   UnclearedBalance = 0
End Function

Private Sub flxBank_Click()
   Dim i As Integer ', iFlxSPayCol As Integer

   With flxBank
      If .TextMatrix(.row, 1) = "" Then Exit Sub

      For i = 1 To .Rows - 1
         .TextMatrix(i, 0) = ""
      Next i
      .TextMatrix(.row, 0) = "X"

      If .TextMatrix(.row, 1) = "" Then Exit Sub

      If .col = 2 Then
         If .TextMatrix(.row, 7) = "" Then
            LoadTypes "Cr"
         Else
            LoadTypes "Dr"
         End If
         picType.Top = .CellTop + .Top - 10
         picType.Left = .CellLeft + .Left - 10
         picType.Width = .ColWidth(2)
         picType.Height = .RowHeight(.row)
         cboType.Top = -10
         cboType.Left = -20
         cboType.Width = picType.Width + 20
         cboType.Height = picType.Height + 10
         cboType.text = .TextMatrix(.row, 2)
         picType.Visible = True
         .ScrollBars = flexScrollBarNone
         picType.ZOrder 0
         picType.SetFocus
         iCurRow = .row
      End If
      If .col > 2 And .col < 6 Then
         If .TextMatrix(.row, 2) = "" Then
            ShowMsgInTaskBar "Please select the type first.", "Y", "N"
            Exit Sub
         End If
      End If
      If .col = 3 Then                                   'Account No
         If picSupList.Visible = True Then Exit Sub
         cmdAccount.Top = .CellTop + .Top - 10
         cmdAccount.Left = .CellLeft + .CellWidth + .Left - cmdAccount.Width - 10
         cmdAccount.Height = .RowHeight(.row)
         cmdAccount.Visible = True
         .ScrollBars = flexScrollBarNone
         cmdAccount.ZOrder 0
         cmdAccount.SetFocus
         iCurRow = .row
      End If
      If .col = 4 Then                                                  'Fund
         If cmdFund.Visible And Not cmdFund.Enabled Then
            If cmdFund.Top = flxBank.CellTop + .Top - 10 Then
               ShowMsgInTaskBar "You donot need to set fund.", "Y", "P"
               Exit Sub
            Else
               cmdFund.Visible = False
            End If
         End If
         If picFund.Visible = True Then Exit Sub
         cmdFund.Top = .CellTop + .Top - 10
         cmdFund.Left = .CellLeft + .CellWidth + .Left - cmdFund.Width - 10
         cmdFund.Height = .RowHeight(.row)
         cmdFund.Visible = True
         .ScrollBars = flexScrollBarNone
         cmdFund.ZOrder 0
         iFundRow = flxBank.row
         If cmdFund.Enabled Then cmdFund.SetFocus
         iCurRow = .row
      End If
      If .col = 5 Then                                                     'Description
         If InStr(.TextMatrix(.row, 2), "Bank") > 0 Then
            ShowMsgInTaskBar "You can add description only when it is Bank Payment or Receipt.", "Y", "N"
            Exit Sub
         End If

         txtDescription.text = ""
         txtDescription.Top = .CellTop + .Top - 10
         txtDescription.Left = .CellLeft + .Left - 10
         txtDescription.Width = .ColWidth(5) - 10
         txtDescription.Height = .RowHeight(.row) - 10
         txtDescription.Visible = True
         .ScrollBars = flexScrollBarNone
         txtDescription.ZOrder 0
         txtDescription.SetFocus
         iCurRow = .row
      End If
   End With
End Sub

Private Sub cboType_Click()
   If flxBank.row <> iCurRow Then Exit Sub

   flxBank.TextMatrix(iCurRow, 2) = cboType.text
   flxBank.TextMatrix(iCurRow, 3) = ""                'After selecting the account no, if user changes
                                                      'the transaction type then user has to select ac no again
   picType.Visible = False
End Sub

Private Sub cboType_LostFocus()
   cboType_Click
   picType.Visible = False
End Sub

Private Sub Form_Load()
   bLoadedSavedTransactions = False

   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString

   LoadClients adoConn
   LoadTypes
   LoadFlxFundList adoConn

'  Load all tenants with balance in gridTenantLookup
   TenantAccountBalance adoConn

   adoConn.Close
   Set adoConn = Nothing

   Me.Height = 8295
   Me.Width = 12885
   Me.Top = 0
   Me.Left = 0

   ConfigFlxBank
   ReDim bSortingCol(flxBank.Cols) As Boolean
   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadFlxFundList(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow As Integer

   ConfigFlxFundList

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not adoRst.EOF
      flxFundList.TextMatrix(iRow, 1) = adoRst.Fields.Item(0).Value
      flxFundList.TextMatrix(iRow, 2) = adoRst.Fields.Item(1).Value
      adoRst.MoveNext
      If Not adoRst.EOF Then flxFundList.AddItem ""
      iRow = iRow + 1
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadTypes(Optional szDebit As String)
   Dim Data() As String

   If IsNull(szDebit) Or szDebit = "" Then
      ReDim Data(0, 5) As String
      Data(0, 0) = "Sales Receipts"
      Data(0, 1) = "Sales Refunds"
      Data(0, 2) = "Purchase Payments"
      Data(0, 3) = "Purchase Refunds"
      Data(0, 4) = "Bank Receipts"
      Data(0, 5) = "Bank Payments"
   End If
   If szDebit = "Cr" Then
      ReDim Data(0, 2) As String
      Data(0, 0) = "Sales Refunds"
      Data(0, 1) = "Purchase Payments"
      Data(0, 2) = "Bank Payments"
   End If
   If szDebit = "Dr" Then
      ReDim Data(0, 2) As String
      Data(0, 0) = "Sales Receipts"
      Data(0, 1) = "Purchase Refunds"
      Data(0, 2) = "Bank Receipts"
   End If
   cboType.Clear
   cboType.Column() = Data()
End Sub

Private Sub LoadClients(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   Dim Data() As String

   ReDim Data(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboClientID.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      flxBank.TextMatrix(iCurRow, 5) = txtDescription.text
      txtDescription.Visible = False
Debug.Print flxBank.TextMatrix(iCurRow, 5)
   End If
End Sub

Private Sub txtDescription_LostFocus()
   flxBank.TextMatrix(iCurRow, 5) = txtDescription.text
   txtDescription.Visible = False
End Sub

Public Sub TestingCommand()
   cboClientID.ListIndex = 1
   cboBC.ListIndex = 0
   cmdBrowse_Click
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
