VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReceiptHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tenant Receipt History"
   ClientHeight    =   10755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13620
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReceiptHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRptIDFrom 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9045
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "1"
      Top             =   270
      Width           =   1215
   End
   Begin VB.TextBox txtRptIDTo 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9045
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "99999999"
      Top             =   690
      Width           =   1215
   End
   Begin VB.PictureBox picLeaseList 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   3135
      Left            =   1920
      ScaleHeight     =   3105
      ScaleWidth      =   6345
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
            Height          =   195
            Index           =   4
            Left            =   3000
            TabIndex        =   23
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   465
         End
         Begin MSForms.ComboBox cboSrcProp 
            Height          =   315
            Left            =   3675
            TabIndex        =   21
            Top             =   0
            Width           =   2295
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "4048;556"
            BoundColumn     =   0
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
            Object.Width           =   "1411"
         End
         Begin MSForms.ComboBox cboSrcClient 
            Height          =   315
            Left            =   480
            TabIndex        =   20
            Top             =   0
            Width           =   2415
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "4260;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   8
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1411"
         End
      End
      Begin VB.TextBox txtTenantSearchUnitName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4080
         TabIndex        =   18
         Top             =   300
         Width           =   1965
      End
      Begin VB.TextBox txtTenantSearchName 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   300
         Width           =   2500
      End
      Begin VB.TextBox txtTenantSearchID 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1425
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
         Left            =   6080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   20
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeaseList 
         Height          =   2490
         Left            =   45
         TabIndex        =   24
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4392
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   70
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Name"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   26
         Top             =   70
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   2
         Left            =   4080
         TabIndex        =   25
         Top             =   75
         Width           =   735
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
         Top             =   80
         Width           =   6015
      End
   End
   Begin VB.CommandButton cmdRptHistoryClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   7
      Top             =   7755
      Width           =   1575
   End
   Begin VB.TextBox txtDateTo 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11715
      MaxLength       =   10
      TabIndex        =   6
      Top             =   690
      Width           =   1215
   End
   Begin VB.TextBox txtDateFrom 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11715
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "01/01/1980"
      Top             =   270
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxReceitHistory 
      Height          =   5415
      Left            =   120
      TabIndex        =   28
      Top             =   2180
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   13553358
      ForeColorFixed  =   12632256
      BackColorSel    =   12622095
      ForeColorSel    =   65535
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Records"
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
      Index           =   4
      Left            =   10200
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit ID"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   10
      Left            =   12120
      TabIndex        =   41
      Top             =   1920
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   120
      X2              =   13440
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt ID From"
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
      Index           =   7
      Left            =   7800
      TabIndex        =   40
      Top             =   270
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt ID To"
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
      Index           =   2
      Left            =   7800
      TabIndex        =   39
      Top             =   690
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mtd"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   9
      Left            =   11040
      TabIndex        =   38
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice ID"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   37
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt ID"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   36
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O/S Amount"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   8
      Left            =   10020
      TabIndex        =   35
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Amount"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   7
      Left            =   8760
      TabIndex        =   34
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   6
      Left            =   6120
      TabIndex        =   33
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Date"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   5
      Left            =   5040
      TabIndex        =   32
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   4560
      TabIndex        =   31
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant Name"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   30
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant ID"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   29
      Top             =   1920
      Width           =   810
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0990F&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   1800
      Width           =   13335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date To"
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
      Index           =   6
      Left            =   10710
      TabIndex        =   13
      Top             =   690
      Width           =   585
   End
   Begin MSForms.CommandButton cmdTenantLookup 
      Height          =   255
      Left            =   7005
      TabIndex        =   2
      Top             =   270
      Width           =   255
      Caption         =   """"
      Size            =   "450;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox cboRptClientList 
      Height          =   315
      Left            =   1035
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4683;556"
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
      Object.Width           =   "1411"
   End
   Begin MSForms.ComboBox cboRptPropertyList 
      Height          =   315
      Left            =   1035
      TabIndex        =   1
      Top             =   735
      Width           =   2655
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4683;556"
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
      Object.Width           =   "1411"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
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
      Index           =   3
      Left            =   195
      TabIndex        =   11
      Top             =   735
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
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
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
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
      Left            =   10710
      TabIndex        =   9
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant"
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
      Index           =   1
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   1020
   End
   Begin MSForms.TextBox txtTenantID 
      Height          =   315
      Left            =   4755
      TabIndex        =   12
      Top             =   240
      Width           =   2535
      VariousPropertyBits=   679495711
      BackColor       =   15858158
      Size            =   "4471;556"
      Value           =   "All Tenants"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E4E4E4&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000F&
      Height          =   1095
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   13335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Height          =   1095
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   13335
   End
End
Attribute VB_Name = "frmReceiptHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboRptClientList_Click()
   RptHistoryRefresh
End Sub

Private Sub cboRptPropertyList_Click()
   RptHistoryRefresh
End Sub

Private Sub cmdGridUnitLookup_Click()
   picLeaseList.Visible = False
End Sub

Private Sub cmdRptHistoryClose_Click()
   Form_Unload 0
End Sub

Private Sub RptHistoryRefresh()
   If flxReceitHistory.Rows = 2 And flxReceitHistory.TextMatrix(1, 0) = "" Then Exit Sub

   Dim i As Integer, szTemp() As String, szTenantName As String

   If txtTenantID.text = "All Tenants" Then
      szTenantName = txtTenantID.text
   Else
      szTemp = Split(txtTenantID.text, " \ ")
      szTenantName = szTemp(1)
   End If

'MsgBox flxReceitHistory.RowHeight(1)
   For i = 1 To flxReceitHistory.Rows - 1
      flxReceitHistory.RowHeight(i) = 240
   Next i

   For i = 1 To flxReceitHistory.Rows - 2
      If Val(flxReceitHistory.TextMatrix(i, 0)) < Val(txtRptIDFrom.text) Or _
         Val(flxReceitHistory.TextMatrix(i, 0)) > Val(txtRptIDTo.text) Or _
         (IIf(szTenantName = "All Tenants", False, True) And flxReceitHistory.TextMatrix(i, 3) <> szTenantName) Or _
         CDate(flxReceitHistory.TextMatrix(i, 5)) < CDate(txtDateFrom.text) Or _
         CDate(flxReceitHistory.TextMatrix(i, 5)) > CDate(txtDateTo.text) Or _
         (IIf(cboRptClientList.text = "All Clients", False, True) And flxReceitHistory.TextMatrix(i, 10) <> cboRptClientList.Column(0)) Or _
         (IIf(cboRptPropertyList.text = "All Properties", False, True) And flxReceitHistory.TextMatrix(i, 11) <> cboRptPropertyList.Column(0)) Then

         flxReceitHistory.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub cmdTenantLookup_Click()
   Me.MousePointer = vbHourglass

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   ConfigureFlxLeaseList
   If cboRptClientList.Column(0) = "ALL" And cboRptPropertyList.Column(0) = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If cboRptClientList.Column(0) <> "ALL" And cboRptPropertyList.Column(0) = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & cboRptClientList.Column(0) & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If cboRptClientList.Column(0) = "ALL" And cboRptPropertyList.Column(0) <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = '" & cboRptPropertyList.Column(0) & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If cboRptClientList.Column(0) <> "ALL" And cboRptPropertyList.Column(0) <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & cboRptClientList.Column(0) & "' AND " & _
               "Units.PropertyID = '" & cboRptPropertyList.Column(0) & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   PopulateTenantLookup adoConn, szSQL

   adoConn.Close
   Set adoConn = Nothing

   txtTenantSearchID.text = ""
   txtTenantSearchName.text = ""
   txtTenantSearchUnitName.text = ""
   picLeaseList.Top = txtTenantID.Top + txtTenantID.Height + 20
   picLeaseList.Left = txtTenantID.Left
   picLeaseList.Visible = True
   picLeaseList.ZOrder 0

   Me.MousePointer = vbArrow
End Sub

Private Sub flxLeaseList_Click()
   txtTenantID.text = flxLeaseList.TextMatrix(flxLeaseList.row, 1) & " \ " & flxLeaseList.TextMatrix(flxLeaseList.row, 2)

   RptHistoryRefresh

   picLeaseList.Visible = False
End Sub

Private Sub Form_Load()
   frmDemands3.Hide
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 8670
   Me.Width = 13710
   Me.BackColor = MODULEBACKCOLOR
   txtDateTo.text = Format(Date, "dd/mm/yyyy")

   ConfigureFlxReceiptHistory

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   If Not g_bFormLoaded Then
      If cboRptClientList.ListCount < 1 Then
         PrepareList adoConn, cboRptClientList, cboRptPropertyList
      End If
   End If

   LoadDataFlxReceitHistory adoConn

   adoConn.Close
   Set adoConn = Nothing

   g_bFormLoaded = True
   Call WheelHook(Me.hWnd)
End Sub

 Public Sub LoadDataFlxReceitHistory(ByVal adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT tlbReceipt.TransactionID, DR.DemandRef, tlbReceipt.SageAccountNumber,  " & _
               "tlbReceipt.RDate, tlbReceipt.Details, tlbReceipt.Amount, tlbReceipt.OSAmount,  " & _
               "SecondaryCode.Value AS PaymentMtd, RIGHT(tlbTransactionTypes.CONSTANT, 2) AS TYPE, " & _
               "Tenants.Name, Property.ClientID, Property.PropertyID, tlbReceipt.UnitID " & _
           "FROM tlbReceipt, (SELECT RptTransactions.FromTran, tlbReceipt.DemandRef " & _
                             "From RptTransactions, tlbReceipt " & _
                             "WHERE RptTransactions.toTran = tlbReceipt.TransactionID) AS DR, " & _
               "tlbTransactionTypes, SecondaryCode, Tenants, LeaseDetails, Units, Property "
   szSQL = szSQL & _
           "WHERE (tlbReceipt.Type = 3 or tlbReceipt.Type = 4)   AND " & _
               "tlbReceipt.Type = tlbTransactionTypes.TYPE_ID AND " & _
               "tlbReceipt.RptAmtType = SecondaryCode.Code AND " & _
               "SecondaryCode.PrimaryCode = 'RAT' AND " & _
               "tlbReceipt.SageAccountNumber = Tenants.SageAccountNumber AND " & _
               "tlbReceipt.TransactionID = DR.FromTran AND " & _
               "Tenants.SageAccountNumber  = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = True AND " & _
               "Units.PropertyID = Property.PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   While Not adoRst.EOF
      flxReceitHistory.TextMatrix(i, 0) = adoRst!TransactionID
      flxReceitHistory.TextMatrix(i, 1) = adoRst!DemandRef
      flxReceitHistory.TextMatrix(i, 2) = adoRst!SageAccountNumber
      flxReceitHistory.TextMatrix(i, 3) = adoRst!Name
      flxReceitHistory.TextMatrix(i, 4) = adoRst!Type
      flxReceitHistory.TextMatrix(i, 5) = adoRst!RDate
      flxReceitHistory.TextMatrix(i, 6) = adoRst!Details
      flxReceitHistory.TextMatrix(i, 7) = Format(adoRst!amount, "0.00")
      flxReceitHistory.TextMatrix(i, 8) = Format(adoRst!OSAmount, "0.00")
      flxReceitHistory.TextMatrix(i, 9) = adoRst!PaymentMtd
      flxReceitHistory.TextMatrix(i, 10) = adoRst!clientID
      flxReceitHistory.TextMatrix(i, 11) = adoRst!propertyID
      flxReceitHistory.TextMatrix(i, 12) = adoRst!unitid

      If Not adoRst.EOF Then flxReceitHistory.AddItem ""
      adoRst.MoveNext
      i = i + 1
   Wend

   Label1(4).Caption = "Total Rec: " & i
   adoRst.Close
   Set adoRst = Nothing
 End Sub

Public Sub ConfigureFlxReceiptHistory()
   Dim szHeader As String, i As Integer

   flxReceitHistory.Clear
   flxReceitHistory.Cols = 13
   flxReceitHistory.Rows = 2

   szHeader$ = "<Receipt ID|<Invoice ID|<Tenant A/C|<Tenant Name|<Type" & _
               "|<Receipt Date|<Details|>Amount £|>O/S Amt. £" & _
               "|<Payment Method|<ClientID|<PropID|<UnitID"
   flxReceitHistory.FormatString = szHeader$

   For i = 0 To flxReceitHistory.Cols - 4
      flxReceitHistory.ColWidth(i) = Label2(i + 1).Left - Label2(i).Left
   Next i
   flxReceitHistory.ColWidth(i + 0) = 0
   flxReceitHistory.ColWidth(i + 1) = 0
   flxReceitHistory.ColWidth(i + 2) = flxReceitHistory.Width - Label2(i).Left - 140

   flxReceitHistory.RowHeight(0) = 0
End Sub

Public Function PopulateTenantLookup(adoConn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Dim adoRst As New ADODB.Recordset

   adoRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   iRow = 1

   While Not adoRst.EOF
      flxLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
      flxLeaseList.TextMatrix(iRow, 2) = adoRst!Name
      flxLeaseList.TextMatrix(iRow, 3) = adoRst!UnitNumber

      iRow = iRow + 1
      adoRst.MoveNext

      If Not adoRst.EOF Then flxLeaseList.AddItem ""
   Wend
   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub ConfigureFlxLeaseList()
   Dim szHeader As String

   flxLeaseList.Clear
   flxLeaseList.Cols = 4
   flxLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Tenant ID|<Tenant Name|<Unit Name"
   flxLeaseList.FormatString = szHeader$
   flxLeaseList.ColWidth(0) = Label20(0).Left - flxLeaseList.Left   '240        Solid column
   flxLeaseList.ColWidth(1) = Label20(1).Left - Label20(0).Left - 20  '1400       'Tenant ID
   flxLeaseList.ColWidth(2) = Label20(2).Left - Label20(1).Left - 20         'Tenant Name
   flxLeaseList.ColWidth(3) = flxLeaseList.Left + flxLeaseList.Width - Label20(2).Left - 300 'Unit Name
   flxLeaseList.Rows = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If g_bFormLoaded Then
      frmDemands3.Show
      Me.Hide
      Cancel = 1
   End If
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   
   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   adoRst.Close

'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   ConfigureFlxLeaseList

'   szSQL = "SELECT Tenants.SageAccountNumber, Name, UnitNumber " & _
'          "From Tenants, LeaseDetails " & _
'          "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
'            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber " & _
'         "ORDER BY Tenants.SageAccountNumber;"
'
'   PopulateTenantLookup adoConn, szSQL

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing

   ConfigureFlxLeaseList

'   szSQL = "SELECT Tenants.SageAccountNumber, Name, UnitNumber " & _
'          "From Tenants, LeaseDetails " & _
'          "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
'            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber " & _
'         "ORDER BY Tenants.SageAccountNumber;"
'
'   PopulateTenantLookup adoConn, szSQL
End Sub

Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_GotFocus()
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   TextBoxFormatDate txtDateFrom
End Sub

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub

Private Sub txtDateTo_LostFocus()
   TextBoxFormatDate txtDateTo
End Sub

Private Sub txtRptIDFrom_GotFocus()
   SelTxtInCtrl txtRptIDFrom
End Sub

Private Sub txtRptIDTo_GotFocus()
   SelTxtInCtrl txtRptIDTo
End Sub

Private Sub txtTenantSearchID_Change()
   Dim i As Integer

   If Len(txtTenantSearchID.text) > 0 Then
      txtTenantSearchName.text = ""
      txtTenantSearchUnitName.text = ""
   End If

   For i = 1 To flxLeaseList.Rows - 1
      flxLeaseList.RowHeight(i) = 240
      If UCase(Left(flxLeaseList.TextMatrix(i, 1), Len(txtTenantSearchID.text))) <> UCase(txtTenantSearchID.text) Then
         flxLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtTenantSearchName_Change()
   Dim i As Integer

   If Len(txtTenantSearchName.text) > 0 Then
      txtTenantSearchID.text = ""
      txtTenantSearchUnitName.text = ""
   End If

   For i = 1 To flxLeaseList.Rows - 1
      flxLeaseList.RowHeight(i) = 240
      If UCase(Left(flxLeaseList.TextMatrix(i, 2), Len(txtTenantSearchName.text))) <> UCase(txtTenantSearchName.text) Then
         flxLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtTenantSearchUnitName_Change()
   Dim i As Integer

   If Len(txtTenantSearchUnitName.text) > 0 Then
      txtTenantSearchID.text = ""
      txtTenantSearchName.text = ""
   End If

   For i = 1 To flxLeaseList.Rows - 1
      flxLeaseList.RowHeight(i) = 240
      If UCase(Left(flxLeaseList.TextMatrix(i, 3), Len(txtTenantSearchUnitName.text))) <> UCase(txtTenantSearchUnitName.text) Then
         flxLeaseList.RowHeight(i) = 0
      End If
   Next i
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
