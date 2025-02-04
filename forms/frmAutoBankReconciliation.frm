VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAutoBankReconciliation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank Reconciliation - Automatic"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutoBankReconciliation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAcBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Height          =   285
      Left            =   10155
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   75
      Width           =   1200
   End
   Begin VB.Frame fraAddTran 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   12000
      TabIndex        =   47
      Top             =   1560
      Width           =   1815
      Begin VB.TextBox txtCloseControl 
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdTenantReceipt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tenant Receipts"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdSupplierPayment 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Supplier Payments"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   675
         Width           =   1575
      End
      Begin VB.CommandButton cmdBRP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bank Receipts && Payments"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1245
         Width           =   1575
      End
      Begin VB.CommandButton cmdBankTransfer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bank Transfers"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSDS 
      Caption         =   "Select a &Different Statement"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   8445
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox txtBankSt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select &Bank Statement"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4870
      Width           =   1815
   End
   Begin VB.CommandButton cmdUnReconTran 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dis&play"
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearTransactions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select another &Bank Account"
      Height          =   405
      Left            =   120
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4020
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   10200
      TabIndex        =   11
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddTran 
      Caption         =   "Add &Transaction"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4020
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear S&election"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   8445
      Width           =   1280
   End
   Begin VB.CommandButton cmdReconcile 
      Caption         =   "Complete &Reconciliation"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   9120
      Width           =   2175
   End
   Begin VB.CommandButton cmdMatched 
      Caption         =   "&Match"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox txtAccountName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Height          =   285
      Left            =   7755
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   80
      Width           =   1200
   End
   Begin VB.CommandButton cmdAutoBankRec 
      Caption         =   "Run &Automatic Bank Reconciliation"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   9120
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBank 
      Height          =   2745
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   4842
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxUnReconTran 
      Height          =   2450
      Left            =   120
      TabIndex        =   21
      Top             =   1545
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   4313
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   6
      Left            =   9480
      TabIndex        =   49
      Top             =   75
      Width           =   600
   End
   Begin VB.Label lblPCB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7320
      TabIndex        =   42
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label lblCB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10080
      TabIndex        =   41
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label lblLSD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6240
      TabIndex        =   40
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   11400
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   11400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Statement Date:"
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   39
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Statement Transactions:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   38
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unreconciled Cashbook Transactions:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   36
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblSOB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10080
      TabIndex        =   34
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unreconciled Total:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   5
      Left            =   5475
      TabIndex        =   33
      Top             =   4050
      Width           =   1455
   End
   Begin VB.Label lblUnclearedBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7080
      TabIndex        =   32
      Top             =   4050
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   4425
      Index           =   3
      Left            =   75
      Top             =   4500
      Width           =   11340
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   4425
      Index           =   2
      Left            =   75
      Top             =   45
      Width           =   11340
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Balance:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   9
      Left            =   8805
      TabIndex        =   31
      Top             =   8540
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Projected Closing Balance:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   8
      Left            =   5445
      TabIndex        =   30
      Top             =   8540
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Opening Balance:"
      ForeColor       =   &H80000007&
      Height          =   300
      Index           =   7
      Left            =   8040
      TabIndex        =   29
      Top             =   4920
      Width           =   2040
   End
   Begin VB.Label Label3 
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
      Height          =   270
      Index           =   5
      Left            =   9600
      TabIndex        =   27
      Top             =   1275
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   4425
      Index           =   1
      Left            =   75
      Top             =   4500
      Width           =   11340
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   4425
      Index           =   0
      Left            =   75
      Top             =   45
      Width           =   11340
   End
   Begin VB.Label Label3 
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
      Height          =   270
      Index           =   4
      Left            =   8160
      TabIndex        =   26
      Top             =   1275
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type"
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   25
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   24
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   23
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   22
      Left            =   7200
      TabIndex        =   20
      Top             =   75
      Width           =   495
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
      Height          =   270
      Index           =   3
      Left            =   8160
      TabIndex        =   18
      Top             =   5355
      Width           =   735
   End
   Begin MSForms.ComboBox cboBC 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   75
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
      Caption         =   "Bank:"
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   17
      Top             =   75
      Width           =   495
   End
   Begin MSForms.ComboBox cboClientID 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   75
      Width           =   3090
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5450;556"
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
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   80
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Matched Ref."
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   15
      Top             =   5385
      Width           =   975
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
      Height          =   270
      Index           =   4
      Left            =   9600
      TabIndex        =   14
      Top             =   5355
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   13
      Top             =   5385
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   5385
      Width           =   375
   End
End
Attribute VB_Name = "frmAutoBankReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szAllBankBalance As String
Private the_array() As String
Private whole_file As String
Private bLoadedSavedTransactions As Boolean

Private Sub cmdAddTran_Click()
   fraAddTran.Left = 2640
   fraAddTran.Top = 1590

   EnaDisFormNotFram "Disable"
   txtCloseControl.SetFocus
End Sub

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

Private Sub cmdAutoBankRec_Click()
   Dim i As Integer, j As Integer

   flxBank.col = 0
   For i = 1 To flxUnReconTran.Rows - 1
      For j = 1 To flxBank.Rows - 1
         flxBank.row = j
         If flxBank.CellBackColor <> RGB(233, 232, 155) And _
         InStr(flxBank.TextMatrix(j, 2), UCase(flxUnReconTran.TextMatrix(i, 3))) > 0 Then
            If flxBank.TextMatrix(j, 4) <> "" Then
               If flxBank.TextMatrix(j, 4) = flxUnReconTran.TextMatrix(i, 6) Then
                  flxBank.TextMatrix(j, 3) = flxUnReconTran.TextMatrix(i, 1)
                  lblCB.Caption = Format(CCur(lblCB.Caption) - CCur(flxBank.TextMatrix(j, 4)), "0.00")
                  HighLightRowsFlxGrid flxBank, j
                  HighLightRowsFlxGrid flxUnReconTran, i
               End If
            End If
            If flxBank.TextMatrix(j, 5) <> "" Then
               If flxBank.TextMatrix(j, 5) = flxUnReconTran.TextMatrix(i, 5) Then
                  flxBank.TextMatrix(j, 3) = flxUnReconTran.TextMatrix(i, 1)
                  lblCB.Caption = Format(CCur(lblCB.Caption) + CCur(flxBank.TextMatrix(j, 5)), "0.00")
                  HighLightRowsFlxGrid flxBank, j
                  HighLightRowsFlxGrid flxUnReconTran, i
               End If
            End If
         End If
      Next j
   Next i

   ShowMsgInTaskBar "Automatic Bank Reconciliation process has been completed."
End Sub

Private Sub cmdBankTransfer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And fraAddTran.Left = 2640 Then
      EnaDisFormNotFram "Enable"
      fraAddTran.Left = 12240
   End If
End Sub

Private Sub cmdBankTransfer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdTenantReceipt.BackColor = &HC0C0C0
   cmdSupplierPayment.BackColor = &HC0C0C0
   cmdBRP.BackColor = &HC0C0C0
   cmdBankTransfer.BackColor = &HF0F0F0
End Sub

Private Sub cmdBRP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And fraAddTran.Left = 2640 Then
      EnaDisFormNotFram "Enable"
      fraAddTran.Left = 12240
   End If
End Sub

Private Sub cmdBRP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdTenantReceipt.BackColor = &HC0C0C0
   cmdSupplierPayment.BackColor = &HC0C0C0
   cmdBRP.BackColor = &HF0F0F0
   cmdBankTransfer.BackColor = &HC0C0C0
End Sub

Private Sub cmdClear_Click()
   If bLoadedSavedTransactions Then
      If MsgBox("It will clear the saved transactions. Do you want to proceed?", vbQuestion + vbYesNo, "Bank Reconciliation") = vbYes Then
         ConfigFlxBank
         cmdBrowse.Enabled = True
         txtBankSt.text = ""
         lblLSD.Caption = ""
         lblSOB.Caption = ""
         lblPCB.Caption = ""
         lblCB.Caption = ""

         Dim adoConn As New ADODB.Connection

         adoConn.Open getConnectionString
         adoConn.Execute "DELETE * FROM tblBankStatement;"
         adoConn.Close
         Set adoConn = Nothing
      End If
   End If
End Sub

Private Sub cmdClearTransactions_Click()
   If Not cmdBrowse.Enabled Then
      MsgBox "There is a reconciliation currently in progress." & Chr(13) & _
             "Please complete or clear this reconciliation before selecting another Bank account to reconcile.", vbCritical + vbOKOnly, "Bank Reconciliation"
      Exit Sub
   End If
   cboClientID.Locked = False
   cboClientID.ListIndex = -1
   cboBC.Locked = False
   cboBC.ListIndex = -1
   ConfigFlxUnReconTran
   lblUnclearedBal.Caption = ""
   txtAccountName.text = ""
   txtAcBal.text = ""
   lblSOB.Caption = ""
   lblPCB.Caption = ""
   lblCB.Caption = ""
   bLoadedSavedTransactions = False
End Sub

Private Sub cmdReconcile_Click()
   If Val(lblPCB.Caption) <> Val(lblCB.Caption) Then
      ShowMsgInTaskBar "Projected Closing Balance is not equal to Closing Balance.", , "N"
      Exit Sub
   End If
'   UnRecon = "X|<SlNumber|<RDate|<ExtRef|<Tran DESCRIPTION|>Dr|>Cr|7"
'   Bank = "X|<Date|<Reference|<Matched Ref.|>Dr|>Cr|>6|7"

   Dim i As Integer
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   flxUnReconTran.col = 1
   For i = 1 To flxUnReconTran.Rows - 1
      flxUnReconTran.row = i
      If flxUnReconTran.CellBackColor = RGB(233, 232, 155) Then
         If flxUnReconTran.TextMatrix(i, 7) = "A" Then            'tlbReceipt
            szSQL = "UPDATE tlbReceipt " & _
                    "SET Reconciled = Amount, ReconNow = '" & Format(Now, "dd/mm/yyyy") & "#Full' " & _
                    "WHERE SlNumber = " & Mid(flxUnReconTran.TextMatrix(i, 1), 3, Len(flxUnReconTran.TextMatrix(i, 1)) - 2) & ";"

            adoConn.Execute szSQL
         End If
         If flxUnReconTran.TextMatrix(i, 7) = "B" Then            'tlbPayment
            szSQL = "UPDATE tlbPayment " & _
                    "SET Reconciled = Amount, ReconNow = '" & Format(Now, "dd/mm/yyyy") & "#Full' " & _
                    "WHERE SlNumber = " & Mid(flxUnReconTran.TextMatrix(i, 1), 3, Len(flxUnReconTran.TextMatrix(i, 1)) - 2) & ";"

            adoConn.Execute szSQL
         End If
         If flxUnReconTran.TextMatrix(i, 7) = "C" Then            'tlbBankPayment
            szSQL = "UPDATE tlbBankPayment " & _
                    "SET Reconciled = Amount, ReconNow = '" & Format(Now, "dd/mm/yyyy") & "#Full' " & _
                    "WHERE TRAN_ID = '" & Mid(flxUnReconTran.TextMatrix(i, 1), 3, Len(flxUnReconTran.TextMatrix(i, 1)) - 2) & "';"

            adoConn.Execute szSQL
         End If
      End If
   Next i

   If bLoadedSavedTransactions Then adoConn.Execute "DELETE * FROM tblBankStatement;"

   ConfigFlxUnReconTran
   ConfigFlxBank
   cboClientID.Locked = False
   cboClientID.ListIndex = -1
   cboBC.Locked = False
   cboBC.ListIndex = -1
   lblUnclearedBal.Caption = ""
   txtAccountName.text = ""
   txtAcBal.text = ""
   lblSOB.Caption = ""
   lblPCB.Caption = ""
   lblCB.Caption = ""
   bLoadedSavedTransactions = False
   cmdBrowse.Enabled = True
   txtBankSt.text = ""
   lblLSD.Caption = ""

   adoConn.Close
   Set adoConn = Nothing
   ShowMsgInTaskBar "Bank Reconciliation has been saved."
End Sub

Private Sub cmdSupplierPayment_Click()
   Load frmPurchaseExpense
   frmPurchaseExpense.tabPurExp.Tab = 1
   frmPurchaseExpense.tabPayment.Tab = 0
   frmPurchaseExpense.Show
   frmPurchaseExpense.ZOrder 0
End Sub

Private Sub cmdSupplierPayment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And fraAddTran.Left = 2640 Then
      EnaDisFormNotFram "Enable"
      fraAddTran.Left = 12240
   End If
End Sub

Private Sub cmdTenantReceipt_Click()
   Load frmDemands3
   frmDemands3.tabDmdRcpt.Tab = 2
   frmDemands3.tabPayment.Tab = 0
   frmDemands3.Show
   frmDemands3.ZOrder 0
End Sub

Private Sub cmdBankTransfer_Click()
   Load frmDemands3
   frmDemands3.tabDmdRcpt.Tab = 2
   frmDemands3.tabPayment.Tab = 2
   frmDemands3.Show
   frmDemands3.ZOrder 0
End Sub

Private Sub cmdBRP_Click()
   Load frmDemands3
   frmDemands3.tabDmdRcpt.Tab = 2
   frmDemands3.tabPayment.Tab = 1
   frmDemands3.Show
   frmDemands3.ZOrder 0
End Sub

Private Sub cmdSupplierPayment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdTenantReceipt.BackColor = &HC0C0C0
   cmdSupplierPayment.BackColor = &HF0F0F0
   cmdBRP.BackColor = &HC0C0C0
   cmdBankTransfer.BackColor = &HC0C0C0
End Sub

Private Sub cmdTenantReceipt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And fraAddTran.Left = 2640 Then
      EnaDisFormNotFram "Enable"
      fraAddTran.Left = 12240
   End If
End Sub

Private Sub cmdTenantReceipt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdTenantReceipt.BackColor = &HF0F0F0
   cmdSupplierPayment.BackColor = &HC0C0C0
   cmdBRP.BackColor = &HC0C0C0
   cmdBankTransfer.BackColor = &HC0C0C0
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
   ConfigFlxUnReconTran
   LoadFlxUnReconTran adoConn
   CheckSavedRecon adoConn

   lblUnclearedBal.Caption = Format(UnclearedBalance, "0.00")

   txtAcBal.text = Format(BankAccBalance(adoConn, cboBC.Column(0), cboClientID.Column(0)), "0.00")

   If Not bConn Then
      adoConn.Close
      Set adoConn = Nothing
   End If
End Sub

Private Sub cmdUnReconTran_Click()
   If cboClientID.text = "" Then
      cboClientID.SetFocus
      Exit Sub
   End If
   If cboBC.text = "" Then
      cboBC.SetFocus
      Exit Sub
   End If
   If txtAccountName.text = "" Then Exit Sub

   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString

   LoadFlxUnReconTran adoConn
   CheckSavedRecon adoConn

   adoConn.Close
   Set adoConn = Nothing

   lblUnclearedBal.Caption = Format(UnclearedBalance, "0.00")

   cboClientID.Locked = True
   cboBC.Locked = True
End Sub

Private Sub txtCloseControl_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And fraAddTran.Left = 2640 Then
      EnaDisFormNotFram "Enable"
      fraAddTran.Left = 12240
   End If
End Sub

Private Sub txtBankSt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Right(txtBankSt.text, 4) = ".csv" Then LoadFileInGrid
   End If
End Sub

Private Sub cboBC_Click()
   If cboBC.text <> "" Then
      txtAccountName.text = cboBC.Column(0)
      lblSOB.Caption = Format(cboBC.Column(4), "0.00")

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
   If cboBC.ListIndex >= 0 Then Exit Sub
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
   If flxBank.TextMatrix(1, 1) = "" Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim i As Integer

'   connect to database
   adoConn.Open getConnectionString

   adoConn.Execute "DELETE * FROM tblBankStatement;"

   szSQL = "SELECT * FROM tblBankStatement;"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'   szHeader$ = "X|<Date|<Reference|<Matched Ref.|>Dr|>Cr|>6|7"
   With adoRst
      For i = 1 To flxBank.Rows - 1
         If flxBank.TextMatrix(i, 1) = "" Then Exit For

         .AddNew
         .Fields.Item("TranDate").Value = Format(flxBank.TextMatrix(i, 1), "dd mmmm yyyy")
         .Fields.Item("ClientBankID").Value = cboBC.Column(2)
         .Fields.Item("StatementReference").Value = flxBank.TextMatrix(i, 2)
         .Fields.Item("MatchedRef").Value = IIf(flxBank.TextMatrix(i, 3) = "", "NM", flxBank.TextMatrix(i, 3)) 'NM -> Not Matched
         .Fields.Item("Dr").Value = IIf(flxBank.TextMatrix(i, 4) = "", 0, flxBank.TextMatrix(i, 4))
         .Fields.Item("Cr").Value = IIf(flxBank.TextMatrix(i, 5) = "", 0, flxBank.TextMatrix(i, 5))
         .Update
      Next i
      .Close
   End With
   
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Bank Reconciliation has been saved."
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
   txtBankSt.text = SelectBankStatement

   If txtBankSt.text = "" Then Exit Sub

   Dim file_name As String
   Dim fnum As Integer
    ' Load the file.
   fnum = FreeFile

   Open txtBankSt.text For Input As fnum
   whole_file = Input$(LOF(fnum), #fnum)
   Close fnum

   LoadFileInGrid
   lblLSD.Caption = Format(flxBank.TextMatrix(1, 1), "dd/mm/yyyy")
   lblPCB.Caption = Format(Val(lblSOB.Caption) + StGridTotal, "0.00")
   lblCB.Caption = Format(Val(lblSOB.Caption), "0.00")
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
'   szHeader$ = "X|<Date|<Reference|>Dr|>Cr|>5|>6|7"
   For r = 1 To num_rows
      flxBank.TextMatrix(r, 1) = the_array(r - 1, 0)
      flxBank.TextMatrix(r, 2) = the_array(r - 1, 1)
      If Val(the_array(r - 1, 2)) < 0 Then
         flxBank.TextMatrix(r, 4) = Format(the_array(r - 1, 2) * (-1), "0.00")
      Else
         flxBank.TextMatrix(r, 5) = Format(the_array(r - 1, 2), "0.00")
      End If
      If r < num_rows Then flxBank.AddItem ""
   Next r

   ShowMsgInTaskBar "Bank statement has been uploaded in the bottom grid successfully"
End Sub

Private Sub CheckSavedRecon(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim r As Integer, i As Integer

   szSQL = "SELECT * FROM tblBankStatement WHERE ClientBankID = " & cboBC.Column(2) & ";"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   szHeader$ = "X|<Date|<Reference|<Matched Ref.|>Dr|>Cr|>6|7"

   If Not adoRst.EOF Then
      bLoadedSavedTransactions = True
      cmdBrowse.Enabled = False
      flxBank.Clear
      flxBank.Rows = 2
      r = 1
      
'      flxBank.AddItem ""
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
            For i = 1 To flxUnReconTran.Rows - 1
               If flxUnReconTran.TextMatrix(i, 1) = flxBank.TextMatrix(r, 3) Then
                  HighLightRowsFlxGrid flxUnReconTran, i
                  Exit For
               End If
            Next i
         End If
         r = r + 1
         adoRst.MoveNext
         If Not adoRst.EOF Then flxBank.AddItem ""
      Wend
      lblPCB.Caption = Format(Val(lblSOB.Caption) + StGridTotal, "0.00")
   Else
      cmdBrowse.Enabled = True
      bLoadedSavedTransactions = False
   End If
   
   adoRst.Close
   Set adoRst = Nothing

   lblLSD.Caption = Format(flxBank.TextMatrix(1, 1), "dd/mm/yyyy")
End Sub

Private Sub ConfigFlxUnReconTran()
   Dim i As Integer
   Dim szHeader As String

   flxUnReconTran.Clear
   
   szHeader$ = "X|<SlNumber|<RDate|<ExtRef|<Tran DESCRIPTION|>Dr|>Cr|7"
   flxUnReconTran.FormatString = szHeader$
   
   flxUnReconTran.Rows = 2
   flxUnReconTran.Cols = 8
   flxUnReconTran.RowHeight(0) = 0

   flxUnReconTran.ColWidth(0) = Label3(0).Left - flxUnReconTran.Left

   For i = 1 To flxUnReconTran.Cols - 3
      flxUnReconTran.ColWidth(i) = Label3(i).Left - Label3(i - 1).Left
   Next i
   flxUnReconTran.ColWidth(i) = flxUnReconTran.Left + flxUnReconTran.Width - Label3(i - 1).Left - 380
   flxUnReconTran.ColWidth(flxUnReconTran.Cols - 1) = 0


   Label3(4).Width = flxUnReconTran.ColWidth(5)
   Label3(5).Width = flxUnReconTran.ColWidth(6) + 50
   cmdUnReconTran.Left = Label3(5).Left
   cmdUnReconTran.Width = Label3(5).Width
End Sub

Private Sub ConfigFlxBank()
   Dim i As Integer
   Dim szHeader As String

   flxBank.Clear
   szHeader$ = "X|<Date|<Reference|<Matched Ref.|>Dr|>Cr|>6|7"
   flxBank.FormatString = szHeader$

   flxBank.Rows = 2
   flxBank.Cols = 8
   flxBank.RowHeight(0) = 0

   flxBank.ColWidth(0) = Label2(0).Left - flxBank.Left

   For i = 1 To flxBank.Cols - 4
      flxBank.ColWidth(i) = Label2(i).Left - Label2(i - 1).Left
   Next i
   flxBank.ColWidth(i) = flxBank.Left + flxBank.Width - Label2(i - 1).Left - 380
   flxBank.ColWidth(flxBank.Cols - 2) = 0
   flxBank.ColWidth(flxBank.Cols - 1) = 0

   Label2(3).Width = Label3(4).Width
   Label2(4).Width = Label3(5).Width
End Sub

Private Sub cmdMatched_Click()
   If flxUnReconTran.TextMatrix(flxUnReconTran.row, 0) = "" Then
      ShowMsgInTaskBar "Please select a transaction from the top grid.", , "N"
      Exit Sub
   End If
   If flxBank.TextMatrix(flxBank.row, 0) = "" Then
      ShowMsgInTaskBar "Please select a transaction from the bottom grid.", , "N"
      Exit Sub
   End If

   If flxBank.TextMatrix(flxBank.row, 4) <> "" Then
      If flxBank.TextMatrix(flxBank.row, 4) = flxUnReconTran.TextMatrix(flxUnReconTran.row, 6) Then
         flxBank.TextMatrix(flxBank.row, 3) = flxUnReconTran.TextMatrix(flxUnReconTran.row, 1)
         lblCB.Caption = Format(CCur(lblCB.Caption) - CCur(flxBank.TextMatrix(flxBank.row, 4)), "0.00")
      Else
         ShowMsgInTaskBar "Payment Amount does not match with Prestige transaction.", , "N"
         Exit Sub
      End If
   End If
   If flxBank.TextMatrix(flxBank.row, 5) <> "" Then
      If flxBank.TextMatrix(flxBank.row, 5) = flxUnReconTran.TextMatrix(flxUnReconTran.row, 5) Then
         flxBank.TextMatrix(flxBank.row, 3) = flxUnReconTran.TextMatrix(flxUnReconTran.row, 1)
         lblCB.Caption = Format(CCur(lblCB.Caption) + CCur(flxBank.TextMatrix(flxBank.row, 5)), "0.00")
      Else
         ShowMsgInTaskBar "Payment Amount does not match with Prestige transaction.", , "N"
         Exit Sub
      End If
   End If

   HighLightRowsFlxGrid flxBank, flxBank.row
   HighLightRowsFlxGrid flxUnReconTran, flxUnReconTran.row

   flxBank.TextMatrix(flxBank.row, 0) = ""
   flxUnReconTran.TextMatrix(flxUnReconTran.row, 0) = ""
End Sub

Private Function UnclearedBalance() As Currency
   Dim iRow As Integer
   On Error Resume Next

   UnclearedBalance = 0

   For iRow = 1 To flxUnReconTran.Rows - 1
      If flxUnReconTran.RowHeight(iRow) > 0 Then
         UnclearedBalance = UnclearedBalance + CCur(flxUnReconTran.TextMatrix(iRow, 5))
         UnclearedBalance = UnclearedBalance - CCur(flxUnReconTran.TextMatrix(iRow, 6))
      End If
   Next iRow
End Function

Private Sub flxBank_Click()
   Dim i As Integer

   With flxBank
      If .TextMatrix(.row, 1) = "" Then Exit Sub

      For i = 1 To .Rows - 1
         .TextMatrix(i, 0) = ""
      Next i
      .TextMatrix(.row, 0) = "X"
   End With
End Sub

Private Sub flxUnReconTran_Click()
   Dim i As Integer

   With flxUnReconTran
      If .TextMatrix(.row, 1) = "" Then Exit Sub

      For i = 1 To .Rows - 1
         flxUnReconTran.TextMatrix(i, 0) = ""
      Next i
      .TextMatrix(.row, 0) = "X"
   End With
End Sub

Private Sub Form_Load()
   bLoadedSavedTransactions = False

   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString

   LoadClients adoConn

   adoConn.Close
   Set adoConn = Nothing

   Me.Height = 10050
   Me.Width = 11565
   Me.BackColor = MODULEBACKCOLOR
   fraAddTran.Left = 12360

   ConfigFlxUnReconTran
   ConfigFlxBank
End Sub

Public Sub LoadFlxUnReconTran(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, j As Integer, iHeaderRow As Integer
   Dim adoSI As New ADODB.Recordset
   Dim szaTran() As String

   On Error GoTo ErrorHandler

   szSQL = "SELECT R.RDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & R.SlNumber AS TID, " & _
               "T.DESCRIPTION AS TT, R.ExtRef AS REF, R.Amount AS AMT, R.Reconciled, " & _
               "R.ReconNow, R.TransactionID, R.Details, R.SageAccountNumber AS ACN, 'A' AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS T " & _
           "WHERE R.BankCode = '" & Trim(cboBC.Column(0)) & "' AND " & _
               "R.Type = T.TYPE_ID AND R.Amount > 0 AND (ISNULL(R.ReconNow) OR R.ReconNow = '') "

   szSQL = szSQL + " UNION "

   szSQL = szSQL + _
           "SELECT P.PDate AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & P.SlNumber AS TID, " & _
               "T.DESCRIPTION AS TT, P.ExtRef AS REF, P.Amount AS AMT, P.Reconciled, " & _
               "P.ReconNow, P.TransactionID, P.Details, P.SageAccountNumber AS ACN, 'B' AS T " & _
           "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
           "WHERE P.BankCode = '" & Trim(cboBC.Column(0)) & "' AND " & _
               "P.Type = T.TYPE_ID AND P.Amount > 0 AND (ISNULL(P.ReconNow) OR P.ReconNow = '') "

   szSQL = szSQL + " UNION "

   szSQL = szSQL + _
           "SELECT BP.TRAN_DATE AS TD, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) & BP.TRAN_ID AS TID, " & _
               "T.DESCRIPTION AS TT, BP.PROJ_REF AS REF, (BP.NET_AMOUNT + BP.VAT) AS AMT, " & _
               "BP.Reconciled, BP.ReconNow, BP.MY_ID AS TransactionID, BP.DESCRIPTION as Details, " & _
               "BP.NOMINAL_CODE AS ACN, 'C' AS T " & _
           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T " & _
           "WHERE BP.BANK_AC = '" & Trim(cboBC.Column(0)) & "' AND " & _
               "BP.TransactionType = T.TYPE_ID AND (BP.NET_AMOUNT + BP.VAT) > 0 AND (ISNULL(BP.ReconNow) OR BP.ReconNow = '') " & _
           "ORDER BY 1;"

'Debug.Print szSQL
   adoSI.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   ConfigFlxUnReconTran

   i = 1

   With flxUnReconTran
      While Not adoSI.EOF
         If Not IsNull(adoSI.Fields.Item("ReconNow")) Then
            szaTran = Split(adoSI.Fields.Item("ReconNow"), "#")
         Else
            ReDim szaTran(1) As String
            szaTran(1) = ""
         End If

         If szaTran(1) = "Full" Then
            .RowHeight(i) = 0
         Else                                            '--------------------------- Part or not Reconciled
            .RowHeight(i) = 240
         End If

         .TextMatrix(i, 1) = adoSI.Fields.Item("TID").Value
         .TextMatrix(i, 2) = adoSI.Fields.Item("TD").Value
         .TextMatrix(i, 3) = IIf(IsNull(adoSI.Fields.Item("REF").Value), "", _
                                                             adoSI.Fields.Item("REF").Value)
         .TextMatrix(i, 4) = adoSI.Fields.Item("TT").Value
         If adoSI.Fields.Item("TT").Value = "Sales Receipt" Or _
            adoSI.Fields.Item("TT").Value = "Sales Receipt on Account" Or _
            adoSI.Fields.Item("TT").Value = "Bank Receipt" Or _
            adoSI.Fields.Item("TT").Value = "Purchase Payment Refund" Then
            .TextMatrix(i, 5) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
         Else
            .TextMatrix(i, 6) = Format(adoSI.Fields.Item("AMT").Value, "0.00")
         End If
         .TextMatrix(i, 7) = adoSI.Fields.Item("T").Value                   'Table indicator. A->Receipt, B->Payment, C->Bank

         i = i + 1

         adoSI.MoveNext
         If Not adoSI.EOF Then .AddItem ""
      Wend
   End With
   adoSI.Close

ErrorHandler:
   Set adoSI = Nothing
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
   cboClientID.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub LoadFlxFund(adoConn As ADODB.Connection)
'   flxFund.Clear
'   flxFund.RowHeight(0) = 0
'   flxFund.Rows = 2
'   flxFund.ColWidth(0) = 1500
'   flxFund.ColWidth(1) = 2700
'   flxFund.ColAlignment = vbLeftJustify
'
'         '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
'   lblSearch0(0).Width = 1400
'   lblSearch0(0).Left = 50
'   lblSearch0(1).Width = 2600
'   lblSearch0(1).Left = lblSearch0(0).Left + flxFund.ColWidth(0)
'
'   txtSearch1.Width = 1400
'   txtSearch1.Left = 40
'   txtSearch2.Width = 2600
'   txtSearch2.Left = txtSearch1.Left + flxFund.ColWidth(0)
'
'   ' Error Handler
'   On Error GoTo Error_Handler
'
'   Dim rRow As Integer, iRec As Integer
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT FundID, FundName FROM Fund;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
'   Else
'      rRow = 1
'      While Not adoRst.EOF
'         flxFund.TextMatrix(rRow, 0) = adoRst.Fields.Item("FundID").Value
'         flxFund.TextMatrix(rRow, 1) = adoRst.Fields.Item("FundName").Value
'         rRow = rRow + 1
'         adoRst.MoveNext
'         If Not adoRst.EOF Then flxFund.AddItem ""
'      Wend
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'   ' Destroy Objects
'   Set adoRst = Nothing
End Sub
'
'Private Sub LoadCboType(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT TYPE_ID, DESCRIPTION " & _
'           "FROM tlbTransactionTypes " & _
'           "WHERE (TYPE_ID > 0 AND TYPE_ID < 13) OR " & _
'               "TYPE_ID = 23 OR TYPE_ID = 24 " & _
'           "ORDER BY TYPE_ID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'      For j = 0 To TotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboType.Column() = Data()
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub
'
'Private Sub txtAccSearchID_Change()
'   Dim i As Integer
'
'   If Len(txtAccSearchID.text) > 0 Then
'      txtAccSearchName.text = ""
'      txtAccTypeSearch.text = ""
'   End If
'
'   For i = 1 To flxAccList.Rows - 1
'      flxAccList.RowHeight(i) = 240
'      If UCase(Left(flxAccList.TextMatrix(i, 1), Len(txtAccSearchID.text))) <> UCase(txtAccSearchID.text) Then
'         flxAccList.RowHeight(i) = 0
'      End If
'   Next i
'End Sub
'
'Private Sub txtAccSearchName_Change()
'   Dim i As Integer
'
'   If Len(txtAccSearchName.text) > 0 Then
'      txtAccSearchID.text = ""
'      txtAccTypeSearch.text = ""
'   End If
'
'   For i = 1 To flxAccList.Rows - 1
'      flxAccList.RowHeight(i) = 240
'      If UCase(Left(flxAccList.TextMatrix(i, 2), Len(txtAccSearchName.text))) <> UCase(txtAccSearchName.text) Then
'         flxAccList.RowHeight(i) = 0
'      End If
'   Next i
'End Sub
'
'Private Sub txtAccTypeSearch_Change()
'   Dim i As Integer
'
'   If Len(txtAccTypeSearch.text) > 0 Then
'      txtAccSearchName.text = ""
'      txtAccSearchID.text = ""
'   End If
'
'   For i = 1 To flxAccList.Rows - 1
'      flxAccList.RowHeight(i) = 240
'      If UCase(Left(flxAccList.TextMatrix(i, 3), Len(txtAccTypeSearch.text))) <> UCase(txtAccTypeSearch.text) Then
'         flxAccList.RowHeight(i) = 0
'      End If
'   Next i
'End Sub
'
'Private Sub cboType_Click()
'   flxBank.TextMatrix(flxBank.row, 9) = cboType.Column(0)
'   flxBank.TextMatrix(flxBank.row, 2) = cboType.Column(1)
'   flxBank.TextMatrix(flxBank.row, 4) = ""
'   picType.Visible = False
'End Sub
'
'Private Sub cboType_LostFocus()
'   picType.Visible = False
'End Sub
'
'Private Sub cmdAccListClose_Click()
'   picAccList.Visible = False
'End Sub
'
'Private Sub cmdGridUnitLookup_Click(Index As Integer)
'   fraList(0).Visible = False
'End Sub
'
'Private Sub flxAccList_Click()
'   picAccList.Visible = False
'   flxBank.SetFocus
'
'   flxBank.TextMatrix(flxBank.row, 4) = flxAccList.TextMatrix(flxAccList.row, 1)
'End Sub
'
'Private Sub flxAccList_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 27 Then
'      picAccList.Visible = False
'      flxBank.SetFocus
'   End If
'End Sub
'
'Private Sub cmdAccLookup_Click()
'   Dim i As Integer
'
'   If flxBank.col = 3 Then
'      fraList(0).Width = 4815
'      cmdAccListClose.Left = fraList(0).Width - cmdAccListClose.Width
'      Shape4(6).Width = fraList(0).Width - cmdAccListClose.Width - 50
'      flxAccList.Width = 4695
'      fraList(0).Left = picAccLookup.Left + 100
'      fraList(0).Top = picAccLookup.Top + 350
'      fraList(0).Visible = True
'      fraList(0).ZOrder 0
''      sTextBox = "Dept"
'      MousePointer = vbDefault
''      flxAccList.SetFocus
'   End If
'
'   If flxBank.col = 4 Then
'      If cboBC.text = "" Then
'         MsgBox "Please select the client's bank account", vbCritical + vbOKOnly, "Bank Account missing"
'         cboBC.SetFocus
'         Exit Sub
'      End If
'
'      With flxAccList
'         For i = 1 To .Rows - 1
'            .RowHeight(i) = 0
'         Next i
'
'         For i = 1 To .Rows - 1
'            If ((Val(flxBank.TextMatrix(flxBank.row, 9)) > 0 And _
'                  Val(flxBank.TextMatrix(flxBank.row, 9)) < 6) Or _
'                  Val(flxBank.TextMatrix(flxBank.row, 9)) = 23) And _
'                  .TextMatrix(i, 4) = "L" Then
'               .RowHeight(i) = 240
'            End If
'            If ((Val(flxBank.TextMatrix(flxBank.row, 9)) > 5 And _
'                  Val(flxBank.TextMatrix(flxBank.row, 9)) < 11) Or _
'                  Val(flxBank.TextMatrix(flxBank.row, 9)) = 24) And _
'                  .TextMatrix(i, 4) = "S" Then
'               .RowHeight(i) = 240
'            End If
'         Next i
'      End With
'
'      txtAccSearchID.text = ""
'      txtAccountName.text = ""
'      txtAccTypeSearch.text = ""
'
'      If picAccLookup.Top + picAccLookup.Height + picAccList.Height < Me.Height Then
'         picAccList.Top = picAccLookup.Top + picAccLookup.Height
'      Else
'         picAccList.Top = picAccLookup.Top - picAccList.Height
'      End If
'
'      picAccList.Left = flxBank.Left + flxBank.CellLeft ' + tabCashbook.Left + tabPayRpt.Left + Frame8(0).Left + txtTenantID.Left + 5
'      picAccList.Visible = True
'      picAccLookup.Visible = False
'      picAccList.ZOrder 0
'      flxAccList.SetFocus
'   End If
'
'   Me.MousePointer = vbDefault
'End Sub

Private Sub LoadFlxAccList(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim iRow As Integer
'
'   szSQL = "SELECT Tenants.SageAccountNumber, Tenants.Name, LeaseDetails.UnitNumber " & _
'           "From Tenants, LeaseDetails " & _
'           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
'            "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
'            "LeaseDetails.Status = True " & _
'          "ORDER BY Tenants.SageAccountNumber;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   iRow = 1
'
'   While Not adoRst.EOF
'      flxAccList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
'      flxAccList.TextMatrix(iRow, 2) = adoRst!Name
'      flxAccList.TextMatrix(iRow, 3) = adoRst!UnitNumber
'      flxAccList.TextMatrix(iRow, 4) = "L"
'
'      iRow = iRow + 1
'      adoRst.MoveNext
'
'      flxAccList.AddItem ""
'   Wend
'   adoRst.Close
'
'   szSQL = "SELECT SupplierID, SupplierName, SupplierType " & _
'           "FROM Supplier " & _
'           "ORDER BY SupplierName;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      flxAccList.TextMatrix(iRow, 1) = adoRst!SupplierID
'      flxAccList.TextMatrix(iRow, 2) = adoRst!SupplierName
'      flxAccList.TextMatrix(iRow, 3) = adoRst!SupplierType
'      flxAccList.TextMatrix(iRow, 4) = "S"
'      adoRst.MoveNext
'      If Not adoRst.EOF Then flxAccList.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'   adoRst.Close
'   Set adoRst = Nothing
End Sub

Private Sub ConfigureFlxAccList()
'   Dim szHeader As String
'
'   flxAccList.Clear
'   flxAccList.Cols = 5
'   flxAccList.RowHeight(0) = 0
'   szHeader$ = "|<ID|<AccName|Type"
'   flxAccList.FormatString = szHeader$
'   flxAccList.ColWidth(0) = Label20(0).Left - flxAccList.Left   '240        Solid column
'   flxAccList.ColWidth(1) = Label20(1).Left - Label20(0).Left - 20  '1400       'Acc ID
'   flxAccList.ColWidth(2) = Label20(2).Left - Label20(1).Left - 20         'Acc Name
'   flxAccList.ColWidth(3) = flxAccList.Left + flxAccList.Width - Label20(2).Left - 300 'Type
'   flxAccList.ColWidth(4) = 0
'   flxAccList.Rows = 2
End Sub
