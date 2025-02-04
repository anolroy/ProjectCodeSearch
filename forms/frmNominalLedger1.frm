VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNominalLedger1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nominal Ledger"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13725
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNominalLedger1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   13725
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6255
      TabIndex        =   38
      Top             =   7080
      Width           =   1125
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   8255
      TabIndex        =   34
      Top             =   7080
      Width           =   1125
   End
   Begin VB.Frame fraLookup 
      BackColor       =   &H00FFEEEE&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   480
      TabIndex        =   21
      Top             =   7560
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdGridTenantLookup 
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
         Height          =   265
         Left            =   4280
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   40
         Width           =   255
      End
      Begin VB.CommandButton cmdSaveLookup 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   340
         Left            =   1740
         TabIndex        =   25
         Top             =   2240
         Width           =   1125
      End
      Begin VB.CommandButton cmdCloseLookup 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   340
         Left            =   3360
         TabIndex        =   24
         Top             =   2240
         Width           =   1125
      End
      Begin VB.CommandButton cmdAddNewLookup 
         Caption         =   "Add &New"
         Height          =   340
         Left            =   120
         TabIndex        =   23
         Top             =   2240
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLookup 
         Height          =   1575
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   2778
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "SAVED"
         Height          =   195
         Index           =   11
         Left            =   2280
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "CLICKED"
         Height          =   195
         Index           =   10
         Left            =   1440
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   630
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   29
         Top             =   285
         Width           =   2205
         VariousPropertyBits=   746604571
         MaxLength       =   100
         BorderStyle     =   1
         Size            =   "3889;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   285
         Width           =   1965
         VariousPropertyBits=   746604571
         MaxLength       =   15
         BorderStyle     =   1
         Size            =   "3466;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Code:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   75
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   26
         Top             =   75
         Width           =   2205
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00400040&
         BorderWidth     =   2
         Height          =   2655
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   2655
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add &New"
      Height          =   375
      Left            =   255
      TabIndex        =   18
      Top             =   7080
      Width           =   1125
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2255
      TabIndex        =   17
      Top             =   7080
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10255
      TabIndex        =   16
      Top             =   7080
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   12255
      TabIndex        =   15
      Top             =   7080
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4255
      TabIndex        =   14
      Top             =   7080
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNominalCode 
      Height          =   5895
      Left            =   255
      TabIndex        =   0
      Top             =   840
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox cboTypeIE 
      Height          =   330
      Left            =   7080
      TabIndex        =   11
      Top             =   465
      Width           =   1920
      VariousPropertyBits=   1753237529
      DisplayStyle    =   3
      Size            =   "3387;582"
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "881;1940"
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Type"
      Height          =   195
      Index           =   14
      Left            =   7200
      TabIndex        =   37
      Top             =   240
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "DrCr"
      Height          =   195
      Index           =   13
      Left            =   9960
      TabIndex        =   36
      Top             =   9000
      Visible         =   0   'False
      Width           =   345
   End
   Begin MSForms.ComboBox cboDrCr 
      Height          =   315
      Left            =   10920
      TabIndex        =   12
      Top             =   480
      Width           =   2160
      VariousPropertyBits=   1753237529
      DisplayStyle    =   3
      Size            =   "3810;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   3
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Dr / Cr"
      Height          =   195
      Index           =   12
      Left            =   11505
      TabIndex        =   35
      Top             =   240
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Type ID"
      Height          =   195
      Index           =   8
      Left            =   9120
      TabIndex        =   20
      Top             =   9000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Category Code"
      Height          =   195
      Index           =   7
      Left            =   7920
      TabIndex        =   19
      Top             =   9000
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSForms.CommandButton cmdLookup 
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   10
      Top             =   510
      Width           =   255
      Caption         =   """"
      Size            =   "450;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdLookup 
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   510
      Width           =   255
      Caption         =   """"
      Size            =   "450;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Debit"
      Height          =   195
      Index           =   5
      Left            =   9000
      TabIndex        =   5
      Top             =   240
      Width           =   1410
   End
   Begin MSForms.TextBox txtName 
      Height          =   315
      Left            =   1335
      TabIndex        =   8
      Top             =   480
      Width           =   2235
      VariousPropertyBits=   679495707
      MaxLength       =   100
      BorderStyle     =   1
      Size            =   "3951;556"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtCode 
      Height          =   315
      Left            =   255
      TabIndex        =   7
      Top             =   480
      Width           =   1080
      VariousPropertyBits=   679495707
      MaxLength       =   15
      BorderStyle     =   1
      Size            =   "1905;556"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Credit"
      Height          =   195
      Index           =   6
      Left            =   10320
      TabIndex        =   6
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Category"
      Height          =   195
      Index           =   4
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Sub Category"
      Height          =   195
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   615
      Index           =   3
      Left            =   120
      Top             =   6960
      Width           =   13455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   615
      Index           =   2
      Left            =   120
      Top             =   6960
      Width           =   13455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000006&
      Height          =   6735
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   13455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   6735
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   13455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Nominal Name"
      Height          =   195
      Index           =   2
      Left            =   1335
      TabIndex        =   2
      Top             =   240
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Code"
      Height          =   195
      Index           =   1
      Left            =   255
      TabIndex        =   1
      Top             =   240
      Width           =   2760
   End
   Begin MSForms.TextBox txtType 
      Height          =   315
      Left            =   5295
      TabIndex        =   13
      Top             =   480
      Width           =   1785
      VariousPropertyBits=   679495711
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "3149;556"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtCategory 
      Height          =   315
      Left            =   3480
      TabIndex        =   32
      Top             =   480
      Width           =   1785
      VariousPropertyBits=   679495711
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "3149;556"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmNominalLedger1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iNewEdit As Byte
Public CALLER_FORM As String

Private Sub cboDrCr_Click()
   Label1(13).Caption = cboDrCr.Column(0)
End Sub

Private Sub cmdAddNew_Click()
   ControlHanlding NewEntryMode
   txtCode.SetFocus
End Sub

Private Sub cmdAddNewLookup_Click()
   TextBox1(0).text = ""
   TextBox1(1).text = ""

   TextBox1(0).SetFocus
   flxLookup.Enabled = False
   cmdAddNewLookup.Enabled = False
   cmdSaveLookup.Enabled = True

   Label1(11).Caption = "NEW"
End Sub

Private Sub cmdCancel_Click()
   If MsgBox("Do you like to discard the changes?", vbQuestion + vbYesNo, "Nominal Ledger") = vbNo Then Exit Sub

   ControlHanlding DefaultMode
   fraLookup.Visible = False
End Sub

Private Sub cmdClose_Click()
   If iNewEdit <> 0 Then
      If MsgBox("Do you want to close without saving the changes?", vbQuestion + vbYesNo, "Nominal Ledger") = vbNo Then Exit Sub
   End If

   Unload Me
End Sub

Private Sub cmdCloseLookup_Click()
   If flxLookup.row = 0 Then
      ShowMsgInTaskBar "Please select a code from the grid.", , "N"
      Exit Sub
   End If

   If Label1(10).Caption = "CATEGORY" Then
      Label1(7).Caption = flxLookup.TextMatrix(flxLookup.row, 0)
      txtCategory.text = flxLookup.TextMatrix(flxLookup.row, 1)
   End If
   If Label1(10).Caption = "TYPE" Then
      Label1(8).Caption = flxLookup.TextMatrix(flxLookup.row, 0)
      txtType.text = flxLookup.TextMatrix(flxLookup.row, 1)
   End If

   TextBox1(0).text = ""
   TextBox1(1).text = ""

   fraLookup.Visible = False
   cmdSaveLookup.Enabled = False
   cmdAddNewLookup.Enabled = True
End Sub

Private Sub cmdDelete_Click()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   
   adoConn.Open getConnectionString

   szSQL = "SELECT NominalCodeforAmount AS NC, 'DemandSplitRecords' AS TN " & _
           "FROM DemandSplitRecords " & _
           "WHERE NominalCodeforAmount = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCodeforAmount AS NC, 'DemandTypes' AS TN " & _
           "FROM DemandTypes " & _
           "WHERE NominalCodeforAmount = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT GlobalBankCode AS NC, 'GlobalData' AS TN " & _
           "FROM GlobalData " & _
           "WHERE GlobalBankCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NC, 'GlobalSCDtls' AS TN " & _
           "FROM GlobalSCDtls " & _
           "WHERE NC = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT PayNCAmt AS NC, 'PayableTypes' AS TN " & _
           "FROM PayableTypes " & _
           "WHERE PayNCAmt = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCode AS NC, 'PayTransactions' AS TN " & _
           "FROM PayTransactions " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCode AS NC, 'RptTransactions' AS TN " & _
           "FROM RptTransactions " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCode AS NC, 'tblPoA' AS TN " & _
           "FROM tblPoA " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NOMINAL_CODE AS NC, 'tblPurInvSRec' AS TN " & _
           "FROM tblPurInvSRec " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCode AS NC, 'TenantDeposit' AS TN " & _
           "FROM TenantDeposit " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NOMINAL_CODE AS NC, 'tlbBankPayment' AS TN " & _
           "FROM tlbBankPayment " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCodeforAmount AS NC, 'tlbChildDemandRecord' AS TN " & _
           "FROM tlbChildDemandRecord " & _
           "WHERE NominalCodeforAmount = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCode AS NC, 'tlbClientBanks' AS TN " & _
           "FROM tlbClientBanks " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NOMINAL_CODE AS NC, 'tlbCreditNote' AS TN " & _
           "FROM tlbCreditNote " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCode AS NC, 'tlbPayment' AS TN " & _
           "FROM tlbPayment " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NominalCode AS NC, 'tlbReceipt' AS TN " & _
           "FROM tlbReceipt " & _
           "WHERE NominalCode = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NOMINAL_CODE AS NC, 'tlbRecharged' AS TN " & _
           "FROM tlbRecharged " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "' "
   szSQL = szSQL + "UNION "

   szSQL = szSQL + "SELECT NOMINAL_CODE AS NC, 'tlbRechargePre' AS TN " & _
           "FROM tlbRechargePre " & _
           "WHERE NOMINAL_CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "';"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount > 0 Then
      MsgBox "The nominal code is being used. You can not delete.", vbCritical + vbOKOnly, "Nominal Code : " & adoRst.Fields.Item(1).Value
   Else
      If MsgBox("Do you wish to delete the nominal code?", vbQuestion + vbYesNo, "Nominal Code") = vbYes Then
         szSQL = "DELETE * " & _
                 "FROM NominalLedger " & _
                 "WHERE Code = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "';"
         adoConn.Execute szSQL

         ShowMsgInTaskBar "Nominal Code " & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & " has been removed."

         szSQL = "SELECT NOMINALLEDGER.Code, NOMINALLEDGER.Name, " & _
                     "NLCategory.CategoryName, NLType.TypeValue, " & _
                     "NOMINALLEDGER.Debit, NOMINALLEDGER.Credit, " & _
                     "NOMINALLEDGER.CategoryCode, NOMINALLEDGER.Type " & _
                 "FROM NOMINALLEDGER, NLCategory, NLType " & _
                 "WHERE NOMINALLEDGER.CategoryCode = NLCategory.CategoryCode AND " & _
                     "NOMINALLEDGER.Type = NLType.NLTypeCode " & _
                 "ORDER BY NOMINALLEDGER.Code;"

         populateGridDefinedHeader adoConn, szSQL, flxNominalCode
      End If
   End If
   
   adoRst.Close
   Set adoRst = Nothing

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdEdit_Click()
   If flxNominalCode.row = 0 Then
      ShowMsgInTaskBar "Please select a Nominal code from the grid.", , "N"
      Exit Sub
   End If

   ControlHanlding EditMode

   txtCode.text = flxNominalCode.TextMatrix(flxNominalCode.row, 0)
   txtName.text = flxNominalCode.TextMatrix(flxNominalCode.row, 1)
   txtCategory.text = flxNominalCode.TextMatrix(flxNominalCode.row, 2)
   txtType.text = flxNominalCode.TextMatrix(flxNominalCode.row, 3)
   Label1(7).Caption = flxNominalCode.TextMatrix(flxNominalCode.row, 7)
   cboDrCr.Value = flxNominalCode.TextMatrix(flxNominalCode.row, 9)
   Label1(8).Caption = flxNominalCode.TextMatrix(flxNominalCode.row, 8)
   cboTypeIE.Value = flxNominalCode.TextMatrix(flxNominalCode.row, 4)
End Sub

Private Sub cmdGridTenantLookup_Click()
   fraLookup.Visible = False
   cmdSaveLookup.Enabled = False
End Sub

Private Sub cmdLookup_Click(Index As Integer)
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   If Index = 0 Then                               'Category Lookup
      szSQL = "SELECT * FROM NLCategory;"
      populateGridDefinedHeader adoConn, szSQL, flxLookup
      Label1(10).Caption = "CATEGORY"
      Label1(0).Caption = "Category"
      fraLookup.Left = txtCategory.Left
      fraLookup.Top = txtCategory.Top
   End If

   If Index = 1 Then                               'Type Lookup
      szSQL = "SELECT * FROM NLType;"
      populateGridDefinedHeader adoConn, szSQL, flxLookup
      Label1(10).Caption = "TYPE"
      Label1(0).Caption = "Sub Category"
      fraLookup.Left = txtType.Left
      fraLookup.Top = txtType.Top
   End If

   flxLookup.row = 0
   flxLookup.col = 0

   adoConn.Close
   Set adoConn = Nothing

   fraLookup.Visible = True
   cmdAddNewLookup.Enabled = True
   cmdAddNewLookup.Enabled = True
   flxLookup.Enabled = True
End Sub

Private Sub cmdPrint_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'  All option selected
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\NL_List.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   
   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdSave_Click()
   If iNewEdit = 0 Then Exit Sub
   If Trim(txtCode.text) = "" Then
      ShowMsgInTaskBar "Please enter the Nominal Code.", , "N"
      txtCode.text = ""
      txtCode.SetFocus
      Exit Sub
   End If
   If Trim(txtName.text) = "" Then
      ShowMsgInTaskBar "Please enter the Name of the Nominal Ledger.", , "N"
      txtName.text = ""
      txtName.SetFocus
      Exit Sub
   End If
   If Trim(txtCategory.text) = "" Then
      ShowMsgInTaskBar "Please select the category of the Nominal Ledger.", , "N"
      txtCategory.text = ""
      txtCategory.SetFocus
      Exit Sub
   End If
   If Trim(txtType.text) = "" Then
      ShowMsgInTaskBar "Please select a sub category of the Nominal Ledger.", , "N"
      txtType.text = ""
      txtType.SetFocus
      Exit Sub
   End If
   If Trim(cboTypeIE.text) = "" Then
      ShowMsgInTaskBar "Please select a type of the Nominal Ledger.", , "N"
      cboTypeIE.text = ""
      cboTypeIE.SetFocus
      Exit Sub
   End If
   If Trim(cboDrCr.text) = "" Then
      ShowMsgInTaskBar "Please select a Dr or Cr of the Nominal Ledger.", , "N"
      cboDrCr.text = ""
      cboDrCr.SetFocus
      Exit Sub
   End If

   On Error GoTo ErrorHandler

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   If iNewEdit = 1 Then
      szSQL = "INSERT INTO NominalLedger (Code, Name, CategoryCode, Type, DrCr, TypeIE) " & _
              "VALUES ('" & txtCode.text & "', '" & txtName.text & "', " & _
                       "'" & Label1(7).Caption & "', '" & Label1(8).Caption & "'," & _
                       "'" & Label1(13).Caption & "', '" & cboTypeIE.Column(0) & "');"
'Debug.Print szSQL
      adoConn.Execute szSQL
'      MsgBox "Nominal Ledger has been added successfully.", vbInformation + vbOKOnly, "Nominal Ledger"
   End If

   If iNewEdit = 2 Then
      szSQL = "UPDATE NominalLedger " & _
              "SET NAME = '" & txtName.text & "', " & _
                  "CategoryCode = '" & Label1(7).Caption & "', " & _
                  "TYPE = '" & Label1(8).Caption & "', " & _
                  "DrCr = '" & Label1(13).Caption & "', " & _
                  "TypeIE = '" & cboTypeIE.Column(0) & "' " & _
              "WHERE CODE = '" & flxNominalCode.TextMatrix(flxNominalCode.row, 0) & "';"
      adoConn.Execute szSQL
'      MsgBox "Nominal Ledger has been edited successfully.", vbInformation + vbOKOnly, "Nominal Ledger"
   End If

   szSQL = "SELECT N.Code, N.Name, " & _
               "NLCategory.CategoryName, NLType.TypeValue, " & _
               "N.TypeIE, N.Debit, N.Credit, " & _
               "N.CategoryCode, N.Type, N.DrCr " & _
           "FROM NOMINALLEDGER AS N, NLCategory, NLType " & _
           "WHERE N.CategoryCode = NLCategory.CategoryCode AND " & _
               "N.Type = NLType.NLTypeCode " & _
           "ORDER BY N.Code;"
'Debug.Print szSQL
   populateGridDefinedHeader adoConn, szSQL, flxNominalCode

   adoConn.Close
   Set adoConn = Nothing

   ControlHanlding DefaultMode

   Exit Sub
ErrorHandler:
   If iNewEdit = 1 Then ShowMsgInTaskBar "System could not add new Nominal Ledger." & Err.description, , "N"
   If iNewEdit = 2 Then ShowMsgInTaskBar "System could not edit Nominal Ledger." & Err.description, , "N"
   ControlHanlding DefaultMode
End Sub

Private Sub cmdSaveLookup_Click()
   If cmdAddNewLookup.Enabled = True And flxLookup.row = 0 Then
      ShowMsgInTaskBar "Please select a code from the grid or cliek Add New.", , "N"
      Exit Sub
   End If
   If cmdAddNewLookup.Enabled = False Then
      If TextBox1(0).text = "" Then
         ShowMsgInTaskBar "Please type the code.", , "N"
         TextBox1(0).SetFocus
         Exit Sub
      End If
      If TextBox1(1).text = "" Then
         ShowMsgInTaskBar "Please type the name.", , "N"
         TextBox1(1).SetFocus
         Exit Sub
      End If
   End If

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, i As Integer

   adoConn.Open getConnectionString

   If Label1(11).Caption = "NEW" Then
      For i = 1 To flxLookup.Rows - 1
         If TextBox1(0).text = flxLookup.TextMatrix(i, 0) Then Exit Sub
      Next i
      If i < flxLookup.Rows - 1 Then
         ShowMsgInTaskBar "The Code already exists.", , "N"
         Exit Sub
      End If

      If Label1(10).Caption = "CATEGORY" Then
         szSQL = "INSERT INTO NLCategory (CategoryCode, CategoryName) " & _
                 "VALUES ('" & TextBox1(0).text & "', '" & TextBox1(1).text & "');"
      End If
      If Label1(10).Caption = "TYPE" Then
         szSQL = "INSERT INTO NLType (NLTypeCode, TypeValue) " & _
                 "VALUES ('" & TextBox1(0).text & "', '" & TextBox1(1).text & "');"
      End If
'Debug.Print szSQL
      adoConn.Execute szSQL
'      MsgBox Label1(10).Caption & " has been added successfully.", vbInformation + vbOKOnly, "Nominal Ledger"
   End If

   If Label1(11).Caption = "SAVED" Then
      If Label1(10).Caption = "CATEGORY" Then
         szSQL = "UPDATE NLCategory " & _
                 "SET CategoryName = '" & TextBox1(1).text & "' " & _
                 "WHERE CategoryCode = '" & TextBox1(0).text & "';"
'         szSQL = "update nlcategory set categoryname='" & TextBox1(1).text & "' where categorycode= " & Val(TextBox1(0).text) & ""
      End If
      If Label1(10).Caption = "TYPE" Then
         szSQL = "UPDATE NLType " & _
                 "SET TypeValue = '" & TextBox1(1).text & "' " & _
                 "WHERE NLTypeCode = '" & Val(TextBox1(0).text) & "';"
      End If
'Debug.Print szSQL
      adoConn.Execute szSQL
'      MsgBox Label1(10).Caption & " has been edited successfully.", vbInformation + vbOKOnly, "Nominal Ledger"
   End If

   flxLookup.Enabled = True
   cmdAddNewLookup.Enabled = True
   cmdSaveLookup.Enabled = False
   Label1(11).Caption = "SAVED"

   If Label1(10).Caption = "CATEGORY" Then                               'Category Lookup
      szSQL = "SELECT * FROM NLCategory;"
      populateGridDefinedHeader adoConn, szSQL, flxLookup
   End If

   If Label1(10).Caption = "TYPE" Then                               'Type Lookup
      szSQL = "SELECT * FROM NLType;"
      populateGridDefinedHeader adoConn, szSQL, flxLookup
   End If

   flxLookup.row = 0
   flxLookup.col = 0

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxLookup_DblClick()
   cmdCloseLookup_Click
End Sub

Private Sub flxLookup_RowColChange()
   TextBox1(0).text = flxLookup.TextMatrix(flxLookup.row, 0)
   TextBox1(1).text = flxLookup.TextMatrix(flxLookup.row, 1)
End Sub

Private Sub flxNominalCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxNominalCode.ToolTipText = flxNominalCode.TextMatrix(flxNominalCode.MouseRow, flxNominalCode.MouseCol)
End Sub

Private Sub Form_Activate()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
   Dim Data(1, 3) As String

   Data(0, 0) = "IN"
   Data(1, 0) = "Income"
   Data(0, 1) = "EX"
   Data(1, 1) = "Expenditure"
   Data(0, 2) = "BS"
   Data(1, 2) = "Balance Sheet"
   Data(0, 3) = "CA"
   Data(1, 3) = "Capital & Reserve"
   cboTypeIE.Column() = Data()

   ConfigureFlxNominalCode
   ControlHanlding DefaultMode

   adoConn.Open getConnectionString
   szSQL = "SELECT N.Code, N.Name, " & _
               "C.CategoryName, NLType.TypeValue, " & _
               "N.TypeIE, N.Debit, N.Credit, " & _
               "N.CategoryCode, N.Type, N.DrCr " & _
           "FROM NOMINALLEDGER AS N, NLCategory AS C, NLType " & _
           "WHERE N.CategoryCode = C.CategoryCode AND " & _
               "N.Type = NLType.NLTypeCode " & _
           "ORDER BY N.Code;"
'Debug.Print szSQL
   populateGridDefinedHeader adoConn, szSQL, flxNominalCode

   szSQL = "SELECT CODE, VALUE " & _
           "FROM   SECONDARYCODE " & _
           "WHERE  PRIMARYCODE = 'NCDC' " & _
           "ORDER BY CODE DESC;"

   populateCombo adoConn, szSQL, cboDrCr

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_Load()
   Me.Height = 8160
   Me.Top = 0
   Me.Left = 0
   Me.BackColor = MODULEBACKCOLOR

   txtCode.Width = Label1(2).Left - Label1(1).Left
   txtName.Width = Label1(3).Left - Label1(2).Left
   txtCategory.Width = Label1(4).Left - Label1(3).Left
   cmdLookup(0).Left = txtCategory.Left + txtCategory.Width - cmdLookup(0).Width - 30
   cmdLookup(0).Top = txtCategory.Top + 30
   txtType.Width = Label1(14).Left - Label1(4).Left
   cmdLookup(1).Left = txtType.Left + txtType.Width - cmdLookup(1).Width - 30
   cmdLookup(1).Top = txtType.Top + 30
   cboDrCr.Left = Label1(12).Left
   cboDrCr.Width = Label1(12).Width
   cboTypeIE.Width = Label1(14).Width
   cboTypeIE.Left = Label1(14).Left

   ConfigureFlxLookup
   flxNominalCode.Sort = 1

   Call WheelHook(Me.hWnd)
End Sub

Private Sub ConfigureFlxLookup()
   flxLookup.Clear
   flxLookup.Cols = 2
   flxLookup.Rows = 2
   flxLookup.RowHeight(0) = 0

   flxLookup.ColWidth(0) = Label1(9).Width
   flxLookup.ColWidth(1) = Label1(0).Width

   fraLookup.Width = 4575
   fraLookup.Height = 2655
End Sub

Private Sub ConfigureFlxNominalCode()
   Dim szHeader As String, iCol As Integer

   flxNominalCode.Clear
   flxNominalCode.Cols = 10
   flxNominalCode.Rows = 2
   flxNominalCode.RowHeight(0) = 0
   szHeader$ = "<Code|<Name|<CategoryCode|<Type|<TypeIE|>Debit|>Credit|>CategoryCode|>Type|<DrCr"
   flxNominalCode.FormatString = szHeader$

   flxNominalCode.ColWidth(0) = Label1(2).Left - Label1(1).Left      'Code
   flxNominalCode.ColWidth(1) = Label1(3).Left - Label1(2).Left      'Nominal Name
   flxNominalCode.ColWidth(2) = Label1(4).Left - Label1(3).Left      'Category
   flxNominalCode.ColWidth(3) = Label1(14).Left - Label1(4).Left     'Type
   flxNominalCode.ColWidth(4) = Label1(5).Left - Label1(14).Left     'Type ID
   flxNominalCode.ColWidth(5) = Label1(6).Left - Label1(5).Left
   flxNominalCode.ColWidth(6) = Label1(12).Left - Label1(6).Left
   flxNominalCode.ColWidth(7) = 0
   flxNominalCode.ColWidth(8) = 0
   flxNominalCode.ColWidth(9) = flxNominalCode.Width + flxNominalCode.Left - Label1(12).Left - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If CALLER_FORM = "frmClientNew4" Then
'      frmClientNew4.LoadNCinCombo
'      frmClientNew4.Show
      CALLER_FORM = ""
   Else
      frmMMain.fraCmdButton.Enabled = True
   End If

   Dim frm As Form

   For Each frm In Forms
     If TypeOf frm Is MDIForm Then
     Else
       ' this is not a mdi form. can use MDIChild property safely
       If frm.MDIChild = True Then
         ' the form is a child of the mdi form
         If frm.Name = "frmGlobal" Then
'            frmGlobal.LoadNCinCombo
         End If
       End If
     End If
   Next frm

   Call WheelUnHook(Me.hWnd)
   Unload Me
End Sub

Private Sub ControlHanlding(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         cmdAddNew.Enabled = True
         cmdEdit.Enabled = True
         cmdSave.Enabled = False
         cmdCancel.Enabled = False
         cmdLookup(0).Enabled = False
         cmdLookup(1).Enabled = False
         cboDrCr.Enabled = False
         cboTypeIE.Enabled = False
         flxNominalCode.Enabled = True

         txtCode.text = ""
         txtCode.Locked = True
         txtName.text = ""
         txtName.Locked = True
         txtCategory.text = ""
         txtCategory.Locked = True
         txtType.text = ""
         txtType.Locked = True

         iNewEdit = 0
   
      Case ComponentMode.NewEntryMode
         cmdAddNew.Enabled = False
         cmdEdit.Enabled = False
         cmdSave.Enabled = True
         cmdCancel.Enabled = True
         cmdLookup(0).Enabled = True
         cmdLookup(1).Enabled = True
         cboDrCr.Enabled = True
         cboTypeIE.Enabled = True
         flxNominalCode.Enabled = False

         iNewEdit = 1
         txtCode.text = ""
         txtCode.Locked = False
         txtName.text = ""
         txtName.Locked = False
         txtCategory.text = ""
         txtType.text = ""
         txtName.SetFocus

      Case ComponentMode.EditMode
         cmdAddNew.Enabled = False
         cmdEdit.Enabled = False
         cmdSave.Enabled = True
         cmdCancel.Enabled = True
         cmdLookup(0).Enabled = True
         cmdLookup(1).Enabled = True
         cboDrCr.Enabled = True
         cboTypeIE.Enabled = True
         flxNominalCode.Enabled = False

         iNewEdit = 2

         txtCode.Locked = True
         txtName.Locked = False
         txtCode.SetFocus

      Case ComponentMode.SavedMode
         cmdAddNew.Enabled = True
         cmdEdit.Enabled = True
         cmdSave.Enabled = False
         cmdCancel.Enabled = False
         cmdLookup(0).Enabled = False
         cmdLookup(1).Enabled = False
         cboDrCr.Enabled = False
         cboTypeIE.Enabled = False
         flxNominalCode.Enabled = True

         txtCode.text = ""
         txtCode.Locked = True
         txtName.text = ""
         txtName.Locked = True
         txtCategory.text = ""
         txtType.text = ""

         iNewEdit = 0
   End Select
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
