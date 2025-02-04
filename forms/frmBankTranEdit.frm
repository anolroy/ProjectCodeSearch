VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBankTranEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankTranEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   13365
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   1215
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   36
      Top             =   6435
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
         TabIndex        =   40
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   39
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   44
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   42
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1620
         TabIndex        =   41
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   37
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   38
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   90
      TabIndex        =   33
      Top             =   4680
      Width           =   13020
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   355
         Left            =   9180
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   270
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   355
         Left            =   10890
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label isOverDratAllowed 
         BackColor       =   &H0080FFFF&
         Height          =   150
         Left            =   945
         TabIndex        =   35
         Top             =   225
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblOverDraftAmount 
         BackColor       =   &H0080FFFF&
         Height          =   150
         Left            =   180
         TabIndex        =   34
         Top             =   225
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4560
      Left            =   90
      TabIndex        =   19
      Top             =   45
      Width           =   13020
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
         Height          =   1320
         Index           =   0
         Left            =   8280
         TabIndex        =   50
         Top             =   3150
         Width           =   4605
         Begin VB.TextBox txtAvailableBankBal1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   2715
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   855
            Width           =   1845
         End
         Begin VB.TextBox txtRetentions1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   2715
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   525
            Width           =   1845
         End
         Begin VB.TextBox txtBankBal1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   2715
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   195
            Width           =   1845
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Avail.Bank Bal£"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   855
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retentions  £"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   525
            Width           =   930
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Balance  £"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   195
            Width           =   1050
         End
      End
      Begin VB.CheckBox chkVat 
         Height          =   285
         Left            =   10620
         TabIndex        =   49
         Top             =   2250
         Width           =   195
      End
      Begin VB.CommandButton cmdVATCode 
         Caption         =   "v"
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
         Left            =   11385
         TabIndex        =   10
         Top             =   2250
         Width           =   255
      End
      Begin VB.CommandButton cmdFund 
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
         Left            =   12555
         TabIndex        =   6
         Top             =   225
         Width           =   300
      End
      Begin VB.CommandButton cmdNC 
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
         Left            =   12555
         TabIndex        =   7
         Top             =   675
         Width           =   300
      End
      Begin VB.CommandButton cmdUnit 
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
         Left            =   6165
         TabIndex        =   3
         Top             =   1710
         Width           =   300
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
         Left            =   6165
         TabIndex        =   2
         Top             =   1215
         Width           =   300
      End
      Begin VB.CommandButton cmdBC 
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
         Left            =   6165
         TabIndex        =   1
         Top             =   720
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
         Left            =   6165
         TabIndex        =   0
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11100
         TabIndex        =   9
         Top             =   1710
         Width           =   1695
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   254
         TabIndex        =   4
         Top             =   2205
         Width           =   5100
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2700
         Width           =   5100
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11100
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   2700
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11100
         TabIndex        =   8
         Top             =   1215
         Width           =   1465
      End
      Begin VB.TextBox txtVat_ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11760
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   2220
         Width           =   1035
      End
      Begin MSForms.TextBox txtBankCode 
         Height          =   285
         Left            =   1350
         TabIndex        =   16
         Top             =   720
         Width           =   1125
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "1984;503"
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
         Caption         =   "T0"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   24
         Left            =   11115
         TabIndex        =   48
         Top             =   2295
         Width           =   180
      End
      Begin MSForms.TextBox txtFund 
         Height          =   285
         Left            =   8910
         TabIndex        =   47
         Top             =   225
         Width           =   3645
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6429;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtNC 
         Height          =   285
         Left            =   8910
         TabIndex        =   46
         Top             =   675
         Width           =   3645
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6429;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtUnit 
         Height          =   285
         Left            =   1350
         TabIndex        =   45
         Top             =   1710
         Width           =   4860
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "8572;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   285
         Left            =   1350
         TabIndex        =   18
         Top             =   1215
         Width           =   4860
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "8572;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBankName 
         Height          =   285
         Left            =   2520
         TabIndex        =   17
         Top             =   720
         Width           =   3690
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6509;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1350
         TabIndex        =   15
         Top             =   225
         Width           =   4860
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "8572;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   32
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account:"
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   31
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   3
         Left            =   8400
         TabIndex        =   30
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net:"
         Height          =   195
         Index           =   10
         Left            =   8400
         TabIndex        =   29
         Top             =   1755
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Rate:"
         Height          =   195
         Index           =   12
         Left            =   8400
         TabIndex        =   28
         Top             =   2205
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   27
         Top             =   2700
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   4
         Left            =   8400
         TabIndex        =   26
         Top             =   2700
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details:"
         Height          =   195
         Index           =   11
         Left            =   270
         TabIndex        =   25
         Top             =   2205
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   24
         Top             =   1215
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Index           =   8
         Left            =   270
         TabIndex        =   23
         Top             =   1710
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "NC:"
         Height          =   195
         Index           =   0
         Left            =   8400
         TabIndex        =   22
         Top             =   675
         Width           =   255
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund:"
         Height          =   195
         Index           =   34
         Left            =   8400
         TabIndex        =   21
         Top             =   270
         Width           =   390
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   285
         Left            =   12570
         TabIndex        =   20
         Top             =   1215
         Width           =   225
         ForeColor       =   8421504
         BackColor       =   16761024
         Caption         =   " P"
         Size            =   "397;503"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmBankTranEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vatOptionEnabled As Boolean
Public FrmBankTranEdit_CALLING_FROM As String             'Name of the form
'Public FrmBankTranEdit_CALLING_MODE As String             'Calling for Add or Edit Bank Transacitons
Private bUnitLoaded  As Boolean
Public szTransID     As String
Dim sTextBox As String

' Private Declare Function MoveWindow& Lib "user32" (ByVal hwnd As Long, _
'      ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
'      ByVal nHeight As Long, ByVal bRepaint As Long)
'
'    Private Sub cboClient_DropDown()
'      SetDropdownHeight cboClient, 1000
'    End Sub
'    Private Sub SetDropdownHeight(cbo As ComboBox, ByVal max_extent As Integer)
'      ' max_extent is the absolute maximum clientY value that the dropdown may extend to
'      ' case 1: nItems <= 8 : do nothing - vb standard behaviour
'      ' case 2: Items will fit in defined max area : resize to fit
'      ' case 3: Items will not fit : resize to defined max height
'
'      If cbo.ListCount > 8 Then
'        Dim max_fit As Integer    ' maximum number of items that will fit in maximum extent
'        Dim item_ht As Integer    ' Calculated height of an item in the dropdown
'
'        item_ht = ScaleY(cbo.Height, ScaleMode, vbPixels) - 8
'        max_fit = (max_extent - cbo.Top - cbo.Height) \ ScaleY(item_ht, vbPixels, ScaleMode)
'
'        If cbo.ListCount <= max_fit Then
'          MoveWindow cbo.hwnd, ScaleX(cbo.Left, ScaleMode, vbPixels), _
'            ScaleY(cbo.Top, ScaleMode, vbPixels), _
'            ScaleX(cbo.Width, ScaleMode, vbPixels), _
'            ScaleY(cbo.Height, ScaleMode, vbPixels) + (item_ht * cbo.ListCount) + 2, 0
'        Else
'          MoveWindow cbo.hwnd, ScaleX(cbo.Left, ScaleMode, vbPixels), _
'            ScaleY(cbo.Top, ScaleMode, vbPixels), _
'            ScaleX(cbo.Width, ScaleMode, vbPixels), _
'            ScaleY(cbo.Height, ScaleMode, vbPixels) + (item_ht * max_fit) + 2, 0
'        End If
'      End If
'    End Sub
'Private Sub cboBC_Change()
'    If cboBC.ListIndex <> -1 Then
'      Label13(7).Caption = cboBC.Column(0)
'      'Modified by anol 26 Apr 2015
'      'It was causing problem on showing overdraft message and function
'      lblOverDraftAmount.Caption = cboBC.Column(3)
'      isOverDratAllowed.Caption = cboBC.Column(2)
'   Else
'      Label13(7).Caption = ""
'      lblOverDraftAmount.Caption = ""
'      isOverDratAllowed.Caption = ""
'   End If
'End Sub



Private Sub chkVat_Click()
    'if this checkbox is not true then do not show vat label and and the selection buttton
    If chkVat.Value = 1 Then
    
        Label1(24).Visible = True
        cmdVATCode.Enabled = True
        txtVat_.Enabled = True
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        Dim szSQL As String
        Dim rstRec As New ADODB.Recordset
        Dim rstVat As New ADODB.Recordset
        Dim nTaxCode As Double
             If txtProperty.text = "" Then
                        'When you do not select a property, take vat code value from the last client
                         'if you open the client form and chage the value you need to instant change the effect. so execute an sql to get the new value.
                         szSQL = "SELECT CLIENTID, CLIENTNAME, CT, V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM ((CLIENT C INNER JOIN Supplier S ON C.ClientID=S.SupplierID) " & _
                        "LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where CLIENTID='" & txtClientList.Tag & "' ORDER BY CLIENTID;"
                         rstVat.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                         If Not rstVat.EOF Then
                                 
                                 Label1(24).Tag = IIf(IsNull(rstVat.Fields("VAT_ID").Value), "-1", rstVat.Fields("VAT_ID").Value)
                                 nTaxCode = IIf(IsNull(rstVat.Fields("VAT_RATE").Value), "0.00", rstVat.Fields("VAT_RATE").Value)
                                 Label1(24).Caption = IIf(IsNull(rstVat.Fields("VAT_CODE").Value), "", rstVat.Fields("VAT_CODE").Value)
                                 txtVat_.text = Format(Val(txtNet.text) * (Val(nTaxCode) / 100), "0.00")
                                 If Label1(24).Tag = -1 Then
                                        chkVat.Value = 0
                                 Else
                                        chkVat.Value = 1
                                 End If
                         Else
                                chkVat.Value = 0
                         End If
                         txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
                         rstVat.Close
                         Set rstVat = Nothing
                Else
                        vatOptionEnabled = LoadVatOption(adoconn)
                        chkVat.Value = IIf((vatOptionEnabled), 1, 0)
                        szSQL = "SELECT P.PropertyID, P.PropertyName,G.VATRate,V.VAT_Rate as RateValue,V.VAT_CODE as VAT_CODE1,G.VATRate  as  Rate,G.vatOptionEnabled " & _
                                         " from ((Property P INNER JOIN globalData G ON P.PropertyID=G.PropertyID) LEFT JOIN tlbVatCode V ON G.VATRate=V.VAT_ID) " & _
                                         "WHERE P.PropertyID = '" & txtProperty.Tag & "' " & _
                                         "ORDER BY P.PropertyID;"
                        rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                        
                        If vatOptionEnabled = True Then
                            chkVat.Value = 1
                            Label1(24).Caption = IIf(IsNull(rstRec.Fields("VAT_CODE1").Value), "", rstRec.Fields("VAT_CODE1").Value)
                            Label1(24).Tag = IIf(IsNull(rstRec.Fields("Rate").Value), "", rstRec.Fields("Rate").Value) 'VAT_ID
                            txtVat_.text = Format(Val(txtNet.text) * (Val(IIf(IsNull(rstRec.Fields("RateValue").Value), "0", rstRec.Fields("RateValue").Value)) / 100), "0.00")
                            cmdVATCode.Enabled = True
                            txtVat_.Enabled = True
                             
                        Else
                            chkVat.Value = 0
                            txtVat_.text = "0.00"
                            Label1(24).Tag = ""
                            Label1(24).Caption = ""
                            cmdVATCode.Enabled = False
                            txtVat_.Enabled = False
                        End If
                        txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
                        rstRec.Close
                        Set rstRec = Nothing
               End If
    Else
        Label1(24).Visible = False
        cmdVATCode.Enabled = False
        txtVat_.Enabled = False
        txtVat_.text = "0.00"
    End If
End Sub

Private Sub cmdFund_Click()
    picClient.Left = 6500.029
    picClient.Top = 555.299
    sTextBox = "6"
    Call loadflxFund
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub
Private Sub FillCboVatCode(adoConnection As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim SQLStr1 As String

   

   SQLStr1 = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM TLBVATCODE where IN_USE Order By VAT_ID"
   adoRst.Open SQLStr1, adoConnection, adOpenDynamic, adLockPessimistic

   txtSearchClientID.text = ""
   txtSearchClientID.Left = 250
   
   txtSearchClientID.Width = 3200
   txtSearchClientName.Visible = False
   
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1200
   flxClient.ColWidth(2) = 1800
   picClient.Width = 3500
   flxClient.Width = 3300
   cmdPicCLose.Left = 3200
   txtSearchClientID.Left = 145
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "VAT Code"
   lblClientName.Caption = ""
   lblClientID.Width = 1400
   lblClientID.Left = 250
   lblClientName.Width = 3600
   Dim rRow As Integer
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
    
    rRow = 1
    While Not adoRst.EOF
       flxClient.row = 1
       flxClient.TextMatrix(rRow, 0) = "  " & adoRst.Fields.Item("VAT_ID").Value
       flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("VAT_CODE").Value
       flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("VAT_RATE").Value
        flxClient.RowHeight(rRow) = 280
       adoRst.MoveNext
       If Not adoRst.EOF Then flxClient.AddItem ""
       rRow = rRow + 1
    Wend
  
   adoRst.Close
   Set adoRst = Nothing
End Sub
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

       '        Case TypeOf ctl Is PictureBox
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            'Mouse wheel was not responding on picturebox
            'this problem fixed by anol 23 Mar 2016
            Case TypeOf ctl Is PictureBox
                        
                            bHandled = False
                       

        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then FocusControl (ctl)

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
Private Sub loadflxFund()
    Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 50
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new
   Dim rsFundMatrix As New ADODB.Recordset
   adoconn.Open getConnectionString
   rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoconn, adOpenStatic, adLockReadOnly
   If rsFundMatrix("isfundAssign").Value = False Then
        szSQL = "SELECT FundID, FundCode,FundName FROM FUND Order by FundCode;"
   Else
        szSQL = "Select * from fundMatrix where PropertyID='" & txtProperty.Tag & "' and ClientID='" & txtClientList.Tag & "' and isDeleted=false"
   End If
   rsFundMatrix.Close
   
   

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If rstRec.EOF Then
        txtFund.text = ""
        txtFund.Tag = ""
   End If
           
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item("FundID").Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundCode").Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundName").Value
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
Private Sub cmdNC_Click()
    picClient.Left = 6500.029
    picClient.Top = 455.299
    sTextBox = "5"
    Call LoadflxNC  'calling the main account here
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdUnit_Click()
    picClient.Left = 269.029
    picClient.Top = 455.299
    sTextBox = "4"
    LoadflxUnit
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdVatCode_Click()
    picClient.Left = 8910.029
    picClient.Top = 200.299
    sTextBox = "7"
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Call FillCboVatCode(adoconn)
    
    adoconn.Close
    
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub Form_Activate()
        Dim strTemp As String
        txtClientList.ForeColor = vbBlack
        If Trim(txtClientList.Tag) = "" Then Exit Sub
        strTemp = Trim(isControlAccountSet(txtClientList.Tag))
        If Len(strTemp) > 0 Then
            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientList.Tag & _
            vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
            strTemp = ""
            txtClientList.ForeColor = vbRed
            Exit Sub
        End If
End Sub

Private Sub txtBC_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
            FocusControl cmdBC
    End If
End Sub

Private Sub cboClient_Change()
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString

   LoadBankAccountInCombo adoconn
   'Resolved by BOSL
   'Bank Payment and Bank Receipt - Incorrect Filtering
   'issue 533
   LoadNCinCombo adoconn
'   'Resolved by BOSL
'   'Bank Payment and Bank Receipt - Incorrect Filtering
'   'modified by anol 10 Aug 2014
'   LoadUnit adoConn
'   LoadCboProperty adoConn
'   'end of modification
   
    'Resolved by BOSL
    'Issue No: 0000467
    'Added By: Asif. 19 Sep 2014
'   LoadCboProperty adoConn
    '''''''''''''''''''''''''''
    
   adoconn.Close
   Set adoconn = Nothing
   
End Sub

Private Sub cboClient_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
            FocusControl cmdBC
    End If
End Sub

Private Sub cboFund_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtDate
    End If
End Sub

Private Sub txtNC_KeyPress(KeyAscii As MSForms.ReturnInteger)
     If KeyAscii = 13 Then
        FocusControl cmdNC
    End If
End Sub

'Private Sub cboProperty_Change()
''   'Resolved by BOSL
''   'Bank Payment and Bank Receipt - Incorrect Filtering
''   'modified by anol 10 Aug 2014
''    If cboProperty.text <> "" Then
''        Dim adoConn As New ADODB.Connection
''        adoConn.Open getConnectionString
''        LoadUnit adoConn
''    End If
'
'   'Resolved by BOSL
'   'Issue No: 0000467
'   'Added By: Asif. 19 Sep 2014
'
'   Dim adoConn As New ADODB.Connection
'   adoConn.Open getConnectionString
'
'   LoadUnit adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   '''''''''''''''''''''''''''
'End Sub
Private Sub LoadBankAccountInCombo(ByVal adoconn As ADODB.Connection)
   On Error GoTo Error_Handler

   If IsNull(txtClientList.Tag) Then Exit Sub
   If txtClientList.text = "" Then Exit Sub

   Dim adoRst As New ADODB.Recordset
   Dim adoRst1 As New ADODB.Recordset
   Dim szSQL As String, Data() As String, j As Integer
   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer
'AllowOverDraft, OverDraftLimit  clause has been added by anol 18 Mar 2015
   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN, AllowOverDraft, OverDraftLimit,DEFAULT_AC " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = NominalLedger.ClientID AND " & _
               "tlbClientBanks.CLIENT_ID <> '' AND " & _
               "NominalLedger.ClientID = '" & txtClientList.Tag & "' AND DEFAULT_AC=true order by tlbClientBanks.NominalCode;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If adoRst.EOF Then
        txtBankName.text = ""
        txtBankCode.text = "" 'NominalLedger.Code
        lblOverDraftAmount.Caption = ""  'OverDraftLimit cboBC.Column(3)
        isOverDratAllowed.Caption = "" 'AllowOverDraft cboBC.Column(2)
        MsgBox "Please set a default Client Bank Account for: " & txtClientList.Tag & "", vbInformation, "Warning"
        szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN, AllowOverDraft, OverDraftLimit,DEFAULT_AC " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = NominalLedger.ClientID AND " & _
               "tlbClientBanks.CLIENT_ID <> '' AND " & _
               "NominalLedger.ClientID = '" & txtClientList.Tag & "'  order by tlbClientBanks.NominalCode;"
        adoRst1.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Not adoRst1.EOF Then
             txtBankName.text = adoRst1.Fields(1).Value
             txtBankCode.text = adoRst1.Fields(0).Value  'NominalLedger.Code
             lblOverDraftAmount.Caption = IIf(IsNull(adoRst1.Fields(3).Value), 0, IsNull(adoRst1.Fields(3).Value)) 'OverDraftLimit cboBC.Column(3)
             isOverDratAllowed.Caption = IIf(IsNull(adoRst1.Fields(2).Value), "False", IsNull(adoRst1.Fields(2).Value)) 'AllowOverDraft cboBC.Column(2)
        Else
             txtBankName.text = ""
             txtBankCode.text = "" 'NominalLedger.Code
             lblOverDraftAmount.Caption = ""  'OverDraftLimit cboBC.Column(3)
             isOverDratAllowed.Caption = "" 'AllowOverDraft cboBC.Column(2)
        End If
   Else
        txtBankName.text = adoRst.Fields(1).Value
        txtBankCode.text = adoRst.Fields(0).Value  'NominalLedger.Code
        lblOverDraftAmount.Caption = IIf(IsNull(adoRst.Fields(3).Value), 0, IsNull(adoRst.Fields(3).Value)) 'OverDraftLimit cboBC.Column(3)
        isOverDratAllowed.Caption = IIf(IsNull(adoRst.Fields(2).Value), "False", IsNull(adoRst.Fields(2).Value)) 'AllowOverDraft cboBC.Column(2)
   End If
   
'   If adoRst.EOF Then GoTo NoRes
'
'   iTotalRow = adoRst.RecordCount
'   iTotalCol = adoRst.Fields.count
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
'   If cboBC.ListCount > 0 Then
'      cboBC.ListIndex = 0
'   End If
'NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

Error_Handler:
   MsgBox Err.description & "::" & Err.Number

   Set adoRst = Nothing
End Sub

Private Sub cmdBC_Click()
    picClient.Left = 269.029
    picClient.Top = 455.299
    sTextBox = "2"
    LoadflxBank
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    FocusControl cmdClientList
End Sub
'Private Sub cboProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        cboUnit1.SetFocus
'    End If
'End Sub

Private Sub cboUnit1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtDetails
    End If
End Sub

'Private Sub cboVat_Click()
'   'Modified by BOSL
'    'issue 463 manual vat
'    'Anol 20 Aug 2014
'    If txtNet.text = "" Then Exit Sub
'    txtVat_.text = Format(Val(txtNet.text) * (cboVat.text / 100), "0.00")
'    txtTotal.text = Format(Val(txtNet.text) * (1 + (cboVat.text / 100)), "0.00")
'End Sub

Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 255.299
    sTextBox = "1"
    LoadflxClient
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub
Private Sub LoadflxBank()
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
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Bank Nominal Code"
   lblClientName.Caption = "Bank Nominal Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
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
  szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN, AllowOverDraft, OverDraftLimit " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = NominalLedger.ClientID AND " & _
               "tlbClientBanks.CLIENT_ID <> '' AND " & _
               "NominalLedger.ClientID = '" & txtClientList.Tag & "';"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
           rRow = 1
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
Private Sub LoadflxUnit()
    Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Unit No"
   lblClientName.Caption = "Unit Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
'   flxClient.Width = 5175
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
           
'        szSQL = "SELECT PropertyID, PropertyName " & _
'                    "FROM Property " & _
'                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                    "ORDER BY PropertyID;"
        If txtClientList.text <> "" Then
             szSQL = " Client.ClientName='" & txtClientList.text & "'"
        End If
        If txtProperty.text <> "" Then
             szSQL = " Property.PropertyName='" & txtProperty.text & "'"
        End If
        If szSQL <> "" Then
             szSQL = " where" & szSQL
        End If
        szSQL = "SELECT UnitNumber, UnitName FROM (Units INNER JOIN Property ON Units.PropertyID=Property.PropertyID) INNER JOIN client ON client.ClientID=Property.ClientID " & szSQL & ";"
       ' Debug.Print szSQL
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            rRow = 1
            
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
           'rRow = rRow + 1
           flxClient.AddItem ""
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = ""
           flxClient.TextMatrix(rRow, 2) = ""
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           'rRow = 2
   
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub LoadflxProperty()
    Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
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
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            rRow = 1
            
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
        
        flxClient.AddItem ""
        flxClient.TextMatrix(rRow, 0) = ""
        flxClient.TextMatrix(rRow, 1) = ""
        flxClient.TextMatrix(rRow, 2) = ""
        flxClient.RowHeight(rRow) = 280
        flxClient.AddItem ""
'           rRow = 2
           
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub LoadflxNC()
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
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Nominal Code"
   lblClientName.Caption = "Nominal Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
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
  szSQL = "SELECT N.* " & _
      "FROM NominalLedger AS N " & _
      "WHERE N.ClientID = '" & txtClientList.Tag & "' AND " & _
      "Posting AND CAFixed=0 AND CODE NOT IN " & _
      "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientList.Tag & "')" & _
      " ORDER BY N.Code;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("CODE").Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("NAME").Value
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
           
           rRow = 1
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

Private Sub cmdproperty_Click()
    picClient.Left = 269.029
    picClient.Top = 455.299
    sTextBox = "3"
    LoadflxProperty
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Function LoadVatOption(Conn As ADODB.Connection) As Integer
    Dim rsGlobalData As New ADODB.Recordset
    rsGlobalData.Open "Select vatOptionEnabled from Globaldata where PropertyID='" & txtProperty.Tag & "'", Conn, adOpenStatic, adLockReadOnly
    If Not rsGlobalData.EOF Then
            LoadVatOption = IIf(IsNull(rsGlobalData("vatOptionEnabled").Value), 0, rsGlobalData("vatOptionEnabled").Value)
    End If
    rsGlobalData.Close
    Set rsGlobalData = Nothing
End Function
Private Sub flxClient_Click()
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        Dim szSQL As String
        Dim rstRec As New ADODB.Recordset
        Dim rstVat As New ADODB.Recordset
        Dim nTaxCode As Double
        Frame1.Enabled = True
        Frame2.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtFund.text = "" 'You need to clear this fund selection because they can be related to client by fund assingment
                txtFund.Tag = ""
                LoadBankAccountInCombo adoconn
                LoadNCinCombo adoconn
                LoadFirstProperty adoconn
                Call updateBankBalance
                FocusControl cmdBC
                Dim strTemp As String
                txtClientList.ForeColor = vbBlack
                strTemp = isControlAccountSet(txtClientList.Tag)
                If Len(strTemp) > 0 Then
                    MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientList.Tag & _
                    vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
                    strTemp = ""
                    picClient.Visible = False
                    txtClientList.ForeColor = vbRed
                    Exit Sub
                End If
        ElseIf sTextBox = "2" Then
                txtBankCode.text = flxClient.TextMatrix(flxClient.row, 1)
                txtBankName.text = flxClient.TextMatrix(flxClient.row, 2)
                If txtDetails.text = "BANK TRANSFER" Then
                    txtNC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtNC.text = flxClient.TextMatrix(flxClient.row, 2)
                End If
                Call updateBankBalance
                FocusControl cmdProperty
        ElseIf sTextBox = "3" Then
                txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtProperty.text = flxClient.TextMatrix(flxClient.row, 2)
                
                 
                FocusControl cmdUnit
                If txtProperty.text = "" Then
                        cmdUnit.Enabled = False
                        txtUnit.Tag = ""
                        txtUnit.text = ""
                Else
                        cmdUnit.Enabled = True
                End If
                
        ElseIf sTextBox = "4" Then
                txtUnit.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtUnit.text = flxClient.TextMatrix(flxClient.row, 2)
                FocusControl txtDetails
        ElseIf sTextBox = "5" Then
                txtNC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtNC.text = flxClient.TextMatrix(flxClient.row, 2)
'                FocusControl cmdFund
                FocusControl txtDate
         ElseIf sTextBox = "6" Then
                txtFund.Tag = flxClient.TextMatrix(flxClient.row, 0)
                txtFund.text = flxClient.TextMatrix(flxClient.row, 2)
'                FocusControl txtDate
                FocusControl cmdFund
        ElseIf sTextBox = "7" Then
                Label1(24).Caption = flxClient.TextMatrix(flxClient.row, 1)
                Label1(24).Tag = flxClient.TextMatrix(flxClient.row, 2)
                txtVat_.text = Format(Val(txtNet.text) * (Val(flxClient.TextMatrix(flxClient.row, 2)) / 100), "0.00")
                txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
                FocusControl cmdSave
        End If
        
        picClient.Visible = False
        adoconn.Close
        Set adoconn = Nothing
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        picClient.Visible = False
        Frame1.Enabled = True
        Frame2.Enabled = True
         If sTextBox = "1" Then
            FocusControl cmdClientList
         ElseIf sTextBox = "2" Then
            FocusControl cmdBC
         ElseIf sTextBox = "3" Then
            FocusControl cmdProperty
         ElseIf sTextBox = "4" Then
            FocusControl cmdUnit
         
         ElseIf sTextBox = "5" Then
            FocusControl cmdNC
         ElseIf sTextBox = "6" Then
            FocusControl cmdFund
         ElseIf sTextBox = "7" Then
            FocusControl cmdVATCode
         End If
        
    End If
    
    If KeyAscii = 13 Then
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
       
        Frame1.Enabled = True
        Frame2.Enabled = True
            If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                LoadBankAccountInCombo adoconn
                LoadNCinCombo adoconn
                LoadFirstProperty adoconn
                cmdBC.SetFocus
            ElseIf sTextBox = "2" Then
                    txtBankCode.text = flxClient.TextMatrix(flxClient.row, 1)
                    txtBankName.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl cmdProperty
            ElseIf sTextBox = "3" Then
                    txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtProperty.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl cmdUnit
            ElseIf sTextBox = "4" Then
                    txtUnit.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtUnit.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl txtDetails
            ElseIf sTextBox = "5" Then
                    txtNC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtNC.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl cmdFund
             ElseIf sTextBox = "6" Then
                    txtFund.Tag = flxClient.TextMatrix(flxClient.row, 0)
                    txtFund.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl txtDate
            ElseIf sTextBox = "7" Then
                    Label1(24).Caption = flxClient.TextMatrix(flxClient.row, 1)
                    Label1(24).Tag = flxClient.TextMatrix(flxClient.row, 2)
                    txtVat_.text = Format(Val(txtNet.text) * (Val(flxClient.TextMatrix(flxClient.row, 2)) / 100), "0.00")
                    txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
                    FocusControl cmdSave
            End If
        
        picClient.Visible = False
        adoconn.Close
        Set adoconn = Nothing
    End If
End Sub

Private Sub Label13_Click(Index As Integer)
   ' MsgBox Label13(7).Caption
End Sub

'Private Sub txtNC_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        cmdNC.SetFocus
'    End If
'End Sub

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
        If sTextBox <> 7 Then
            FocusControl txtSearchClientName
           End If
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 27 Then
'         picClient.Visible = False
'          Frame1.Enabled = True
'          Frame2.Enabled = True
'          If sTextBox = "1" Then
'                 cmdClientList.SetFocus
''           ElseIf sTextBox = "2" Then
''                cmdproperty.SetFocus
''           ElseIf sTextBox = "3" Then
''                cmdFundLookUp.SetFocus
'           End If
'    End If
If KeyAscii = 27 Then
        picClient.Visible = False
        Frame1.Enabled = True
        Frame2.Enabled = True
         If sTextBox = "1" Then
         ElseIf sTextBox = "1" Then
            FocusControl cmdClientList
         ElseIf sTextBox = "2" Then
            FocusControl cmdBC
         ElseIf sTextBox = "3" Then
            FocusControl cmdProperty
         ElseIf sTextBox = "4" Then
            FocusControl cmdUnit
         
         ElseIf sTextBox = "5" Then
            FocusControl cmdNC
         ElseIf sTextBox = "6" Then
            FocusControl cmdFund
         ElseIf sTextBox = "7" Then
            FocusControl cmdVATCode
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
         FocusControl flxClient
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            FocusControl flxClient
        End If
    End If
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdClientList
    End If
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtReference
    End If
End Sub

Private Sub txtReference_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        FocusControl cmdNC
    End If
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        picClient.Visible = False
        Frame1.Enabled = True
        Frame2.Enabled = True
         If sTextBox = "1" Then
         ElseIf sTextBox = "1" Then
            FocusControl cmdClientList
         ElseIf sTextBox = "2" Then
            FocusControl cmdBC
         ElseIf sTextBox = "3" Then
            FocusControl cmdProperty
         ElseIf sTextBox = "4" Then
            FocusControl cmdUnit
         
         ElseIf sTextBox = "5" Then
            FocusControl cmdNC
         ElseIf sTextBox = "6" Then
            FocusControl cmdFund
         ElseIf sTextBox = "7" Then
            FocusControl cmdVATCode
         End If
        
    End If
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
End Sub



Private Sub txtUnit_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdUnit
    End If
End Sub

Private Sub txtVat__Change()
    'Resolved by BOSL
    'issue 463 manual vat
    'newly added by Anol 20 Aug 2014
    If IsNumeric(txtVat_.text) = False Then
        txtVat_.text = "0.00"
    End If
End Sub

Private Sub txtVat__KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
End Sub

Private Sub txtVat__LostFocus()
    'Resolved by BOSL
    'issue 463 manual vat
    'newly added by Anol 20 Aug 2014
     txtVat_.text = Format(txtVat_.text, "0.00")
     txtTotal.text = Format((Val(txtNet.text) + Val(txtVat_.text)), "0.00")
End Sub
Private Sub cboVat_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
'   If Val(txtTotal.text) = 0 Then
'      ShowMsgInTaskBar "0 value transaction will not be saved.", "Y", "N"
'      SelTxtInCtrl txtTotal
'      Exit Sub
'   End If
    'issue 519
    'it should not be possible to change the posting date on a transaction
    'if the posting date on that transaction falls within a closed financial period.
    
    If txtClientList.ForeColor = vbRed Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientList.text & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        Exit Sub
    End If
    If txtClientList.text = "" Then
      MsgBox "Please select a Client.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl cmdClientList
      
      Exit Sub
   End If
   If txtBankName.text = "" Then
      MsgBox "Please select a Bank Account for this transaction.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl cmdBC
      
      Exit Sub
   End If
'   If txtProperty.text = "" And txtClientList.text = "" Then
'      MsgBox "Please select a Property for this transaction.", vbExclamation + vbOKOnly, "Saving..."
'      FocusControl cmdProperty
'
'      Exit Sub
'   End If
   If Trim(txtProperty.text) = "" Then
          If MsgBox("You have not selected a property. Do you wish to add a property?", vbYesNo, "Select a Property") = vbYes Then
              cmdProperty.SetFocus
              Exit Sub
          End If
   End If
   If txtDetails.text = "" Then
      MsgBox "Please enter all the required information to save this transaction.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl txtDetails
      
      Exit Sub
   End If
   If txtReference.text = "" Then
      MsgBox "Please enter a reference for this transaction.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl txtReference
      
      Exit Sub
   End If
   If txtNC.text = "" Then
      MsgBox "Please select a Nominal Code for this transaction.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl cmdNC
      
      Exit Sub
   End If
   If txtFund.text = "" Then
      MsgBox "Please select a Fund for this transaction.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl cmdFund
      
      Exit Sub
   End If
   If txtDate.text = "" Then
      MsgBox "Please enter the transaction date.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl txtDate
      
      Exit Sub
   End If
'   If txtNet.text = "" Or Val(txtNet.text) = 0 Then
'      MsgBox "Please enter the amount.", vbExclamation + vbOKOnly, "Saving..."
'      txtNet.SetFocus
'
'      Exit Sub
'   End If
   If Label1(24).Caption = "" And chkVat.Value = 1 Then
      MsgBox "Please select a VAT Code for this transaction.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl cmdVATCode
      
      Exit Sub
   End If
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    If IsPeriodStatus(lblPostingDate.ToolTipText, txtClientList.Tag, adoconn) = 0 Then
        ShowMsgInTaskBar "The posting date cannot fall within a closed Financial Period", "Y", "N"
        adoconn.Close
        Exit Sub
    ElseIf IsPeriodStatus(lblPostingDate.ToolTipText, txtClientList.Tag, adoconn) = 9 Then
        ShowMsgInTaskBar "The Posting date does not fall in any existing Financial Period", "Y", "N"
        adoconn.Close
        Exit Sub
    End If
    adoconn.Close
    If DateDiff("d", lblPostingDate.ToolTipText, txtDate.text) > 0 Then
          MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
          Exit Sub
    End If
    cmdSave.Enabled = False
    If Left(Me.Caption, 3) = "Add" Then
        If Not SaveNewTrans Then Exit Sub
        txtClientList.text = ""
        txtClientList.Tag = ""
        txtProperty.text = ""
        txtProperty.Tag = ""
        txtUnit.text = ""
        txtUnit.Tag = ""
        txtBankName.text = ""
        txtBankCode.text = ""
        txtDetails.text = ""
        txtReference.text = ""
'        If IsLoadedAndVisible("frmCashbook") = True Then
'            frmCashbook.cboBC_Click
'        End If
      If MsgBox("Would you like to add another Bank " & Right(Me.Caption, 7) & "?", vbQuestion + vbYesNo) = vbNo Then
         Unload Me
      Else
         txtDetails.text = ""
         txtReference.text = ""
         txtNC.text = ""
         txtNC.Tag = ""
         txtFund.text = ""
         txtNet.text = ""
         cmdSave.Enabled = True
'         Label1(24).Caption = ""
'         Label1(24).Tag = ""
         txtTotal = "0.00"
         'modidfied by anol 20161013
         'cmdClientList.SetFocus
         Call clientFocus
      End If
   End If
   If Left(Me.Caption, 4) = "Edit" Then
        SaveEditTrans
        If IsLoadedAndVisible("frmCashbook") = True Then
            frmCashbook.cboBC_Click
        End If
   End If
End Sub
Private Sub clientFocus()
    On Error GoTo Err
    cmdClientList.SetFocus
    Exit Sub
    
Err:
End Sub
Private Sub UpdateOverDraftStatus(adoconn As ADODB.Connection, szBank As String, szClientID As String)
      Dim rsClientBank As New ADODB.Recordset
      rsClientBank.Open "Select AllowOverDraft,OverDraftLimit from tlbClientBanks where NominalCode='" & szBank & "' and Client_ID='" & szClientID & "'", adoconn, adOpenStatic, adLockReadOnly
      If Not rsClientBank.EOF Then
            isOverDratAllowed.Caption = rsClientBank("AllowOverDraft").Value
            lblOverDraftAmount.Caption = IIf(IsNull(rsClientBank("OverDraftLimit").Value), "0", rsClientBank("OverDraftLimit").Value)
       End If
       rsClientBank.Close
       Set rsClientBank = Nothing
End Sub
Private Function SaveNewTrans() As Boolean
   Dim szStr As String
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   'On Error GoTo ErrHandler
   adoconn.Open getConnectionString
   adoconn.BeginTrans
   '0000525: Bank overdrawn message showing on transaction entry incorrectly
   'Overdraft Checking By anol 18 Mar 2015
   Dim dblBankBanlance As Double
   If Right(Me.Caption, 7) = "Payment" Then
        dblBankBanlance = BankAccBalance(adoconn, txtBankCode.text, txtClientList.Tag)
        'add Retention value effective with Bank balance by anol 2023-05-24
        szSQL = "Select sum(amount) as DAmt from RetentionDetails where  isDeleted=false and BankCode='" & txtBankCode & "' and ClientID='" & txtClientList.Tag & "'"
        adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Not adoRst.EOF Then
            dblBankBanlance = dblBankBanlance - IIf(IsNull(adoRst.Fields.Item("DAmt").Value), 0, adoRst.Fields.Item("DAmt").Value)
        End If
        adoRst.Close
    
        If dblBankBanlance - Val(txtTotal.text) < 0 Then 'Account balance-Current Amount
              Call UpdateOverDraftStatus(adoconn, txtBankCode.text, txtClientList.Tag) ''szBank As String, szClientID As String
              If isOverDratAllowed.Caption = "True" Then
                 If Val(lblOverDraftAmount.Caption) > 0 Then
                    If (dblBankBanlance - Val(txtTotal.text)) * (-1) > Val(lblOverDraftAmount.Caption) Then
                       If MsgBox("This Bank Account is over its overdraft limit. Do you wish to continue?", vbQuestion + vbYesNo, "Bank Overdrawn") = vbNo Then Exit Function
                    End If
                 End If
              Else
                 MsgBox "This Bank Account cannot go overdrawn", vbInformation + vbOKOnly, "Bank Overdraft"
                 Exit Function
              End If
        End If
   End If
'End of modification
  
   szStr = "SELECT * FROM tlbBankPayment;"
'Debug.Print szStr
   adoRst.Open szStr, adoconn, adOpenDynamic, adLockOptimistic
   With adoRst
      .AddNew
      .Fields.Item("MY_ID").Value = UniqueID()
      .Fields.Item("CreatedBy").Value = User
      .Fields.Item("CreatedDate").Value = Now
      .Fields.Item("ClientID").Value = txtClientList.Tag
      .Fields.Item("BANK_AC").Value = txtBankCode.text  'cbobc.value
      .Fields.Item("PropertyID").Value = txtProperty.Tag
      .Fields.Item("UNIT_ID").Value = txtUnit.Tag
      .Fields.Item("DESCRIPTION").Value = txtDetails.text
      .Fields.Item("PROJ_REF").Value = txtReference.text
      .Fields.Item("NOMINAL_CODE").Value = txtNC.Tag
      .Fields.Item("DEPT_ID").Value = txtFund.Tag
      .Fields.Item("TRAN_DATE").Value = txtDate.text
      .Fields.Item("NET_AMOUNT").Value = Val(txtNet.text)
      If chkVat.Value = 1 Then 'added by anol 2020-09-13
            .Fields.Item("TAX_CODE").Value = Label1(24).Caption
            .Fields.Item("VAT").Value = Val(txtTotal.text) - Val(txtNet.text)
      Else
            .Fields.Item("TAX_CODE").Value = Null
            .Fields.Item("VAT").Value = 0
      End If
      .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")

      If Right(Me.Caption, 7) = "Receipt" Then
         .Fields.Item("TransactionType").Value = 12
         .Fields.Item("TRANS").Value = "BR"
         UpdateBankAcBal_Plus adoconn, Val(txtTotal.text), txtBankCode.text, txtClientList.Tag  ''cbobc.value

      End If
      If Right(Me.Caption, 7) = "Payment" Then
         .Fields.Item("TransactionType").Value = 11
         .Fields.Item("TRANS").Value = "BP"
          UpdateBankAcBal_Minus adoconn, Val(txtTotal.text), txtBankCode.text, txtClientList.Tag
      End If
      .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoconn)
      If txtBankCode.text = txtNC.Tag Then       'Contra Transation
         .Fields.Item("CT").Value = "C"
      End If

      .Update
      'Resolved by BOSL
      'issue 546
      'Bank transfer not showing bank balance
'UpdateBankAcBal_Plus adoConn, CCur(flxSPayment.TextMatrix(iRow, 11)), frmBRPreForm.cmbBankAc.Column(3), frmBRPreForm.txtClientlist.tag
'UpdateBankAcBal_Plus adoConn,
      If txtBankCode.text = txtNC.Tag Then       'Contra Transation
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = txtClientList.Tag
         .Fields.Item("BANK_AC").Value = txtBankCode.text
         .Fields.Item("PropertyID").Value = txtProperty.Tag
         .Fields.Item("UNIT_ID").Value = txtUnit.Tag
         .Fields.Item("DESCRIPTION").Value = txtDetails.text
         .Fields.Item("PROJ_REF").Value = txtReference.text
         .Fields.Item("NOMINAL_CODE").Value = txtNC.Tag
         .Fields.Item("DEPT_ID").Value = txtFund.Tag
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtNet.text)
         .Fields.Item("TAX_CODE").Value = Label1(24).Caption
         .Fields.Item("VAT").Value = Val(txtTotal.text) - Val(txtNet.text)
         .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
   
         If Right(Me.Caption, 7) = "Payment" Then
            .Fields.Item("TransactionType").Value = 12
            .Fields.Item("TRANS").Value = "BR"
         End If
         If Right(Me.Caption, 7) = "Receipt" Then
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
         End If
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", adoconn)
         .Fields.Item("CT").Value = "C"
         .Update
      End If

      .Close
   End With

   Set adoRst = Nothing

   If FrmBankTranEdit_CALLING_FROM = "frmBankTransactions" Then
      Call frmBankTransactions.LoadFlxBankPay(adoconn, "")
   End If
    SaveNewTrans = True
'--------------------------------------------------------------------------------------------
'  Export Transactions to Nominal Ledger (NLPosting table)
   If Export_BPnBR_2_NL(adoconn) = True Then
        adoconn.CommitTrans
        ShowMsgInTaskBar "The Transaction has been saved.", "Y", "P"
   Else
        adoconn.RollbackTrans
        MsgBox "There was a problem saving this transaction. It has therefore been rolled back", vbInformation, "Transaction rolled back"
   End If
   
'--------------------------------------------------------------------------------------
   FocusControl cmdClose
   adoconn.Close
   Set adoconn = Nothing

   
   Exit Function
End Function

Private Sub SaveEditTrans()
   Dim szStr As String
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

   On Error GoTo ErrHandler
'      connect to database
   adoconn.Open getConnectionString

   szStr = "SELECT BP.* " & _
           "FROM tlbBankPayment AS BP " & _
           "WHERE MY_ID = '" & szTransID & "'"
'Debug.Print szStr
   adoRst.Open szStr, adoconn, adOpenDynamic, adLockOptimistic
   With adoRst
      If .EOF Then
         GoTo ErrHandler
      Else
         .Fields.Item("ClientID").Value = txtClientList.Tag
         .Fields.Item("BANK_AC").Value = txtBankCode.text
         .Fields.Item("PropertyID").Value = txtProperty.Tag
         .Fields.Item("UNIT_ID").Value = txtUnit.Tag
         .Fields.Item("DESCRIPTION").Value = txtDetails.text
         .Fields.Item("PROJ_REF").Value = txtReference.text
         .Fields.Item("NOMINAL_CODE").Value = txtNC.Tag
         .Fields.Item("DEPT_ID").Value = txtFund.Tag
         .Fields.Item("TRAN_DATE").Value = txtDate.text
         .Fields.Item("NET_AMOUNT").Value = Val(txtNet.text)
         If chkVat.Value = 1 Then 'added by anol 2020-09-13
            .Fields.Item("TAX_CODE").Value = Label1(24).Caption
            .Fields.Item("VAT").Value = Val(txtTotal.text) - Val(txtNet.text)
        Else
              .Fields.Item("TAX_CODE").Value = Null
              .Fields.Item("VAT").Value = 0
        End If
'         .Fields.Item("TAX_CODE").Value = Label1(24).Caption
'         .Fields.Item("VAT").Value = Val(txtTotal.text) - Val(txtNet.text)
         .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
         
         'Resolved By BOSL. Issue: 0000503. Modified by Asif
         .Fields.Item("NLPost").Value = False
         .Fields.Item("LastModifiedBy").Value = User
         .Fields.Item("LastModifiedDate").Value = Now
         .Update
         .Close
      End If
   End With
   
   
'  Resolved By BOSL. Issue: 0000503. Modified by Asif

   DeleteJournalNLPosting adoconn, szTransID
'  Export Transactions to Nominal Ledger (NLPosting table)
   Export_BPnBR_2_NL adoconn
'--------------------------------------------------------------------------------------
   

   ShowMsgInTaskBar "The Transaction has been saved.", "Y", "P"

   Set adoRst = Nothing
   
   If FrmBankTranEdit_CALLING_FROM = "frmBankPaymentHistory" Then
      frmBankPaymentHistory.LoadFlxBankPay adoconn
   End If
   If FrmBankTranEdit_CALLING_FROM = "frmBankTransactions" Then
      Call frmBankTransactions.LoadFlxBankPay(adoconn, "")
   End If
   If FrmBankTranEdit_CALLING_FROM = "frmBankTransactionsHistory" Then
      Call frmBankTransactions.LoadflxBankPayHist(adoconn, "")
   End If

   adoconn.Close
   Set adoconn = Nothing

   Unload Me
   Exit Sub

ErrHandler:
   MsgBox "System could not update the record.", vbExclamation + vbOKOnly, "Edit Bank Transactions"

   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub Form_Load()
   Dim adoconn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   'if this checkbox is not true then do not show vat label and and the selection buttton
    If chkVat.Value = True Then
        Label1(24).Visible = True
        cmdVATCode.Enabled = True
    Else
        Label1(24).Visible = False
        cmdVATCode.Enabled = False
    End If


   Me.ZOrder 0
   Me.Height = 6090
   Me.Width = 13230


   bUnitLoaded = False

        adoconn.Open getConnectionString
        ' Clients
        szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
        adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Not adoRst.EOF Then
                txtClientList.Tag = adoRst.Fields("CLIENTID").Value
                txtClientList.text = adoRst.Fields("CLIENTNAME").Value
                adoRst.Close
                LoadFirstProperty adoconn
                LoadBankAccountInCombo adoconn
                LoadNCinCombo adoconn
               

       End If

   LoadVAT adoconn
   
   adoconn.Close
   Set adoconn = Nothing
   Call updateBankBalance
   Call WheelHook(Me.hWnd)

End Sub
Public Sub updateBankBalance()
        Dim adoconn As New ADODB.Connection
        Dim adoRst As New ADODB.Recordset
        adoconn.Open getConnectionString
        Dim Balance As Double
        Dim szSQL As String
   ' find current Balance for the selected bank account and selected client ID by anol 2023-05-24
   szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) as AMTT from (" & _
            "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND R.BankCode = '" & txtBankCode & "' AND U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND P.ClientID = '" & txtClientList.Tag & "' AND B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID group by Type UNION "
                  
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & txtBankCode & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & txtClientList.Tag & "' AND B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID  group by TRANS UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.BankCode = '" & txtBankCode & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & txtClientList.Tag & "'   group by Type )"
                       
    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
      txtBankBal1.text = IIf(IsNull(adoRst.Fields.Item("AMTT").Value), 0, adoRst.Fields.Item("AMTT").Value)
      txtBankBal1.text = Format(txtBankBal1.text, "0.00")
   End If
   adoRst.Close
    szSQL = "Select sum(amount) as DAmt from RetentionDetails where isDeleted=false and BankCode='" & txtBankCode & "' and ClientID='" & txtClientList.Tag & "'"
    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
        txtRetentions1.text = IIf(IsNull(adoRst.Fields.Item("DAmt").Value), 0, adoRst.Fields.Item("DAmt").Value)   'adoRst.Fields.Item("DAmt").Value
        txtRetentions1.text = Format(txtRetentions1.text, "0.00")
    End If
    adoRst.Close
    
    
'   szSQL = "SELECT * from tlbClientBanks where NominalCode and client_ID='" & txtClientList & "'"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   If Not adoRst.EOF Then
'   End If
   txtAvailableBankBal1.text = Val(txtBankBal1.text) - Val(txtRetentions1.text)
   txtAvailableBankBal1.text = Format(txtAvailableBankBal1.text, "0.00")
'   txtAvailableBankBal1
'   txtRetentions1
   adoconn.Close
End Sub
Private Sub LoadVAT(adoconn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim szaData() As String, i As Integer

   On Error GoTo ErrorHandler
   
   szSQL = "SELECT VAT_CODE, VAT_RATE FROM tlbVATCODE ORDER BY VAT_ID"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'   ReDim szaData(1, adoRst.RecordCount) As String
'
'   i = 0
'   While Not adoRst.EOF
'      szaData(0, i) = adoRst.Fields.Item("VAT_CODE").Value
'      szaData(1, i) = adoRst.Fields.Item("VAT_RATE").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'
'   cboVat.Clear
'   cboVat.Column() = szaData()
    If Not adoRst.EOF Then
         Label1(24).Caption = adoRst.Fields("VAT_CODE").Value
         Label1(24).Tag = adoRst.Fields("VAT_RATE").Value
    End If
   adoRst.Close
   Set adoRst = Nothing
   'cboVat.ListIndex = 0
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

'Private Sub LoadUnit(adoConn As ADODB.Connection)
'    Dim szSQL As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szaData() As String, i As Integer
'   'Resolved by BOSL
'   'issue 451 ;Bank Payment and Bank Receipt - Incorrect Filtering
'   'modified by anol 08 Aug 2014
'   On Error GoTo ErrorHandler
'   If txtClientList.text <> "" Then
'        szSQL = " Client.ClientName='" & txtClientList.text & "'"
'   End If
'   If txtProperty.text <> "" Then
'        szSQL = " Property.PropertyName='" & txtProperty.text & "'"
'   End If
'   If szSQL <> "" Then
'        szSQL = " where" & szSQL
'   End If
'        szSQL = "SELECT UnitNumber, UnitName FROM (Units INNER JOIN Property ON Units.PropertyID=Property.PropertyID) INNER JOIN client ON client.ClientID=Property.ClientID " & szSQL & ";"
'       ' Debug.Print szSQL
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim szaData(1, adoRst.RecordCount) As String
'
'   cboUnit.Clear
'   cboUnit1.Clear
'   i = 0
'   While Not adoRst.EOF
'      szaData(0, i) = adoRst.Fields.Item("UnitNumber").Value
'      szaData(1, i) = adoRst.Fields.Item("UnitName").Value
'      cboUnit1.AddItem adoRst.Fields.Item("UnitName").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'   cboUnit.Column() = szaData()
'   adoRst.Close
'   Set adoRst = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

'Private Sub LoadFund(adoConn As ADODB.Connection)
'   Dim szSQL As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szaData() As String, i As Integer
'
'   On Error GoTo ErrorHandler
'
' 'Resolved by BOSL
' 'Issue No: 0000467
' 'Modified By: Asif. 19 Sep 2014
'
'   'szSQL = "SELECT FundID, FundName FROM FUND;"
'   szSQL = "SELECT FundID, FundCode FROM FUND;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim szaData(1, adoRst.RecordCount) As String
'
'   i = 0
'   While Not adoRst.EOF
'      szaData(0, i) = adoRst.Fields.Item("FundID").Value
'      'szaData(1, i) = adoRst.Fields.Item("FundName").Value
'      szaData(1, i) = adoRst.Fields.Item("FundCode").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'' End of Modification
'
'   cboFund.Clear
'   cboFund.Column() = szaData()
'
'   adoRst.Close
'   Set adoRst = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub LoadNCinCombo(adoconn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, TotalRow As Integer
   Dim Data() As String, i As Integer
   'issue 533
   'added by anol 06 Feb 2015
   If IsNull(txtClientList.Tag) = True Then Exit Sub
 
'issue 502
'modified by anol 17 Dec 2014
 szSQL = "SELECT N.* " & _
      "FROM NominalLedger AS N " & _
      "WHERE N.ClientID = '" & txtClientList.Tag & "' AND " & _
      "Posting AND (ISNULL(CAType) OR CAType='') AND CODE NOT IN " & _
      "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientList.Tag & "')" & _
      " ORDER BY N.Code;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount
   ReDim Data(2, TotalRow) As String

   i = 0
   While Not adoRst.EOF
      Data(0, i) = adoRst.Fields.Item("Code").Value
      Data(1, i) = adoRst.Fields.Item("Name").Value
      i = i + 1
      adoRst.MoveNext
   Wend

   'cboNC.Column() = Data()

   ' Destroy Objects
   Set adoRst = Nothing
End Sub
'Private Sub cboUnit1_Change()
'    'Resolved by BOSL
'   'Bank Payment and Bank Receipt - Incorrect Filtering
'   'modified by anol 11 Aug 2014
'    cboUnit.ListIndex = cboUnit1.ListIndex
'End Sub
'Private Sub cboUnit1_Click()
'  'Resolved by BOSL
'   'Bank Payment and Bank Receipt - Incorrect Filtering
'   'modified by anol 11 Aug 2014
'    cboUnit.ListIndex = cboUnit1.ListIndex
'End Sub
Private Sub LoadFirstProperty(adoconn As ADODB.Connection) 'load first property if there is only one property issue 713
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   
   
'   'Resolved by BOSL
'   'issue 451 ;Bank Payment and Bank Receipt - Incorrect Filtering
'   'modified by anol 08 Aug 2014
''*************************************** PROPERTY ******************************************
'   If txtClientlist.text <> "" Then
'        szSQL = " AND Client.ClientName='" & txtClientlist.text & "'"
'   End If


    'Resolved by BOSL
    'Issue No: 0000467
    'If the Client is selected then load only the properties of the selected clients
    'Modfied By: Asif.19 Sep 2014
   ' cboProperty.Clear
    '*************************************** PROPERTY ******************************************
    If txtClientList.Tag <> "" Then
'        szSQL = "SELECT Property.PropertyID, PropertyName " & _
'               "FROM Property, GlobalData, tlbVatCode " & _
'               "WHERE Property.PropertyID = GlobalData.PropertyID AND " & _
'                   "GlobalData.VATRate = tlbVatCode.VAT_ID " & _
'                   "AND Property.ClientID = '" & txtClientList.Tag & "' " & _
'               "ORDER BY Property.PropertyID;"
                szSQL = "SELECT Property.PropertyID, PropertyName " & _
               "FROM Property where " & _
               " Property.ClientID = '" & txtClientList.Tag & "' " & _
               "ORDER BY Property.PropertyID;"
'    Else
'       szSQL = "SELECT Property.PropertyID, PropertyName " & _
'               "FROM Property, GlobalData, tlbVatCode " & _
'               "WHERE Property.PropertyID = GlobalData.PropertyID AND " & _
'                   "GlobalData.VATRate = tlbVatCode.VAT_ID " & _
'               "ORDER BY Property.PropertyID;"
           adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

           If adoRst.RecordCount = 1 Then
        '      Dim TotalRow As Integer, TotalCol As Integer
        '      Dim Data() As String
        '      Dim i As Integer, j As Integer
        '
        '      TotalRow = adoRst.RecordCount - 1
        '      TotalCol = adoRst.Fields.count - 1
        '
        '      ReDim Data(TotalCol, TotalRow) As String
        '
        '      For i = 0 To TotalRow
        '         For j = 0 To TotalCol
        '            Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
        '         Next j
        '            adoRst.MoveNext
        '         If adoRst.EOF Then Exit For
        '      Next i
        '      cboProperty.Column() = Data()
                txtProperty.text = adoRst.Fields("PropertyName").Value
                txtProperty.Tag = adoRst.Fields("PropertyID").Value
           Else
                 txtProperty.text = ""
                 txtProperty.Tag = ""
           End If
           adoRst.Close
           Set adoRst = Nothing
    End If
    '''''''''''''''''''''''''''''''''''''''''''

         'szSQL = ""


   'Resolved by BOSL
   'Issue No: 0000467
   'Added By: Asif. 19 Sep 2014

'   If cboProperty.ListCount > 0 Then
'        cboProperty.ListIndex = 0
'   End If
   '''''''''''''''''''''''''''''''
   
  
End Sub

'Private Sub LoadCboClient(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
''   On Error GoTo ErrorHandler
''
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
'           "FROM     CLIENT " & _
'           "ORDER BY CLIENTID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      Dim TotalRow As Integer, TotalCol As Integer
'      Dim Data() As String
'      Dim i As Integer, j As Integer
'
'      TotalRow = adoRst.RecordCount - 1
'      TotalCol = adoRst.Fields.count - 1
'
'      ReDim Data(TotalCol, TotalRow) As String
'
'      For i = 0 To TotalRow
'          For j = 0 To TotalCol
'              Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'          Next j
'          adoRst.MoveNext
'          If adoRst.EOF Then Exit For
'      Next i
'
'      cboClient.Column() = Data()
'      cboClient.ListIndex = 0
'   End If
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim adoconn As New ADODB.Connection
   If frmBankTransactions.bEditMode = True Then
          adoconn.Open getConnectionString
        adoconn.Execute "Update tlbBankPayment Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & frmBankTransactions.UserSessionID & "'"
             frmBankTransactions.bEditMode = False
        adoconn.Close
        Set adoconn = Nothing
   End If
   If FrmBankTranEdit_CALLING_FROM = "frmBankPaymentHistory" Then frmBankPaymentHistory.Enabled = True
   If InStr(FrmBankTranEdit_CALLING_FROM, "frmBankTransactions") > 0 Then frmBankTransactions.Enabled = True
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
    'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
        If txtClientList.text = "" Then
            ShowMsgInTaskBar "Please select a Client.", "Y", "N"
            Exit Sub
        End If
'        If frmMMain.IsRibbonVersion Then
'        Dim adoConn As New ADODB.Connection
'        Dim szSQL As String
'        adoConn.Open getConnectionString
'        If IsPeriodStatus(lblPostingDate.ToolTipText, cboClient.Column(0), adoConn) = 0 Then
'           ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
'           adoConn.Close
'           Set adoConn = Nothing
'           Exit Sub
'        ElseIf IsPeriodStatus(lblPostingDate.ToolTipText, cboClient.Column(0), adoConn) = 9 Then
'           ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
'           adoConn.Close
'           Set adoConn = Nothing
'           Exit Sub
'        End If
'    End If
        'End of modification
        If IsDate(lblPostingDate.ToolTipText) = False Then
            ShowMsgInTaskBar "Date format is not correct"
            Exit Sub
        End If
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDate.text, txtClientList.Tag
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
   lblPostingDate.ToolTipText = txtDate.text
End Sub

Private Sub txtDate_GotFocus()
   If Len(txtDate.text) < 10 Then txtDate.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNet.SetFocus
    End If
    TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
  If txtDate.text <> "" Then
      If TextBoxFormatDate(txtDate) Then
         If lblPostingDate.ToolTipText = "" And lblPostingDate.ToolTipText <> txtDate.text Then
            lblPostingDate.ToolTipText = txtDate.text
         End If
      End If
   End If
   'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
        If txtClientList.text = "" Then
            ShowMsgInTaskBar "Please select a Client.", "Y", "N"
            Exit Sub
        End If
        If frmMMain.IsRibbonVersion Then
        Dim adoconn As New ADODB.Connection
        Dim szSQL As String
        'fixed by anol 15 Sep 2015 issue 571 note 1172
        If IsDate(txtDate.text) = False Then Exit Sub
        adoconn.Open getConnectionString
        If IsPeriodStatus(txtDate.text, txtClientList.Tag, adoconn) = 0 Then
           ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtDate.text, txtClientList.Tag, adoconn) = 9 Then
           ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        End If
    End If
        'End of modification
End Sub

Private Sub txtNet_Change()
'   Dim tot As Currency
'
'   On Error GoTo ErrHndl
'
'   If txtNet.text = "" Or cboVat.text = "" Then Exit Sub
'
'   tot = CCur(txtNet.text) + CCur(cboVat.text)
'   txtTotal.text = Format(Val(txtNet.text) * (1 + (cboVat.text / 100)), "0.00")
'
'   Exit Sub
'
'ErrHndl:
'   txtTotal.text = "0.00"
'   txtNet.text = "0.00"
'   txtTotal.text = "0.00"
    'Dim tot As Currency

   On Error GoTo ErrHndl

   If txtNet.text = "" Or Label1(24).Tag = "" Then Exit Sub

  ' tot = CCur(txtNet.text) + CCur(cboVat.text)
   txtTotal.text = Format(Val(txtNet.text) * (1 + (Label1(24).Tag / 100)), "0.00")

   Exit Sub

ErrHndl:
   txtTotal.text = "0.00"
   txtNet.text = "0.00"
   txtTotal.text = "0.00"
End Sub

Private Sub txtNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 10 Then txtVat_.SetFocus
'    If KeyAscii = 13 And cboVat.Enabled Then
'        cboVat.SetFocus
'    End If
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtNet_LostFocus()
''    Dim szTemp() As String
''    'Resolved by BOSL
''    'issue 463:Manually Typing in the VAT Value
''    'Modified by Anol 02 Sep 2014
''
''   If cboVat.text <> "" Then
''      szTemp = Split(cboVat.text, "  ")
''      txtVat_.text = Format(Val(txtNet.text) * (Val(szTemp(0)) / 100), "0.00")
''   End If
''   'cboVat.text = Format(IIf(txtNet.text = "", 0, Val(txtNet.text)) * (nTaxCode / 100), "0.00")
''   txtNet.text = Format(txtNet.text, "0.00")
''   txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
   
   If Label1(24).Tag <> "" Then
      txtVat_.text = Format(Val(txtNet.text) * (Val(Label1(24).Tag) / 100), "0.00")
   End If

   txtNet.text = Format(Val(txtNet.text), "0.00")
   txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
   
End Sub
