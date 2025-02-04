VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBankPaymentHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Receipt & Payment History"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13590
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankPaymentHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   13590
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   12180
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Transaction"
      Height          =   400
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankPay 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   345
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   9446
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
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   255
      Index           =   12
      Left            =   12240
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "N/C"
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      Height          =   255
      Index           =   9
      Left            =   10320
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "T/C"
      Height          =   255
      Index           =   10
      Left            =   11160
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "TAX"
      Height          =   255
      Index           =   11
      Left            =   11520
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblBankRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmBankPaymentHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bFrmCashBookOpen   As Boolean
Private sCallingForm       As Single

Private Sub cmdClose_Click()
   Form_Unload 1
End Sub

Private Sub Form_Activate()
   If bFrmCashBookOpen Then frmCashbook.Hide
End Sub

Private Sub Form_Load()
   If frmDemands3.Visible Then
      frmDemands3.Hide
      sCallingForm = 1
   Else
      frmBankTransactions.Hide
      sCallingForm = 2
   End If

   Me.Height = 6720
   Me.Width = 13650
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR

   'Define Bank Payment flex grid
   ConfigFlxBankGrid

   Dim adoConn As New ADODB.Connection
'      connect to database
   adoConn.Open getConnectionString

   LoadFlxBankPay adoConn

   adoConn.Close
   Set adoConn = Nothing

   BANK_PAYMENT_HISTORY_LOADED = True
   bFrmCashBookOpen = IsLoadedAndVisible("frmCashbook")

   Call WheelHook(Me.hWnd)
End Sub

Private Sub ConfigFlxBankGrid()
   Dim szHeader As String, iCol As Integer

   szHeader$ = "<No|<Bank|<Type|<Date|<UnitID|<N/C|<Ref|<Fund|<Details|>Net|<T/C|>TAX|>Total|ID|Client|PropID|FundID"

   flxBankPay.Clear
   flxBankPay.Rows = 2
   flxBankPay.Cols = 19

   flxBankPay.RowHeight(0) = 0
   flxBankPay.FormatString = szHeader$

   For iCol = 1 To flxBankPay.Cols - 7
      flxBankPay.ColWidth(iCol - 1) = lblBankRec(iCol).Left - lblBankRec(iCol - 1).Left
   Next iCol

   flxBankPay.ColWidth(12) = flxBankPay.Width + flxBankPay.Left - lblBankRec(12).Left - 300
   flxBankPay.ColWidth(13) = 0         'ID field
   flxBankPay.ColWidth(14) = 0         'Client Id
   flxBankPay.ColWidth(15) = 0         'Porperty ID
   flxBankPay.ColWidth(16) = 0         'FundID
   flxBankPay.ColWidth(17) = 0         'Bank Rec
   flxBankPay.ColWidth(18) = 0         'Client Name
End Sub

Public Sub LoadFlxBankPay(adoConn As ADODB.Connection)
'   Dim iRow As Integer
'   Dim szStr As String
'   Dim adoRst As New ADODB.Recordset
'
'   szStr = "SELECT BP.*, F.FundName,C.ClientName,NL.Name " & _
'           "FROM tlbBankPayment AS BP, Fund AS F,Client C,NominalLedger NL " & _
'           "WHERE BP.DEPT_ID = CSTR(F.FundID) AND NL.ClinetID =C.ClientID AND BP.Nominal_code=NL.code AND C.ClientID=BP.ClientID AND " & _
'               "BP.UPDATE_SAGE = TRUE " & _
'           "ORDER BY BP.TRANS, CLNG(BP.TRAN_ID);"
'   adoRst.Open szStr, adoConn, adOpenDynamic, adLockPessimistic
''Debug.Print szStr
''  szHeader$ = "<No|<Bank|<Type|<Date|<UnitID|<N/C|<Ref|<Fund|<Details|>Net|<T/C|>TAX|>Total|ID|UNIT_ID|PropID|FundID"
'
'   flxBankPay.Clear
'   flxBankPay.Rows = 2
'   iRow = 1
'   While Not adoRst.EOF
'      flxBankPay.TextMatrix(iRow, 0) = adoRst!TRANS & adoRst!TRAN_ID
'      flxBankPay.TextMatrix(iRow, 1) = adoRst!BANK_AC
'      flxBankPay.TextMatrix(iRow, 2) = IIf(IsNull(adoRst!TRANS), "", adoRst!TRANS)
'      flxBankPay.TextMatrix(iRow, 3) = adoRst!TRAN_DATE
'      flxBankPay.TextMatrix(iRow, 4) = IIf(IsNull(adoRst!clientID), "", adoRst!clientID)
'      flxBankPay.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!NOMINAL_CODE), "", adoRst!NOMINAL_CODE)
'      flxBankPay.TextMatrix(iRow, 6) = IIf(IsNull(adoRst!PROJ_REF), "", adoRst!PROJ_REF)
'      flxBankPay.TextMatrix(iRow, 7) = IIf(IsNull(adoRst!FundName), "", adoRst!FundName)
'      flxBankPay.TextMatrix(iRow, 8) = IIf(IsNull(adoRst!description), "", adoRst!description)
'      flxBankPay.TextMatrix(iRow, 9) = Format(IIf(IsNull(adoRst!NET_AMOUNT), "", adoRst!NET_AMOUNT), "0.00")
'      flxBankPay.TextMatrix(iRow, 10) = IIf(IsNull(adoRst!TAX_CODE), "", adoRst!TAX_CODE)
'      flxBankPay.TextMatrix(iRow, 11) = Format(IIf(IsNull(adoRst!vat), "0", adoRst!vat), "0.00")
'      flxBankPay.TextMatrix(iRow, 12) = Format(Val(flxBankPay.TextMatrix(iRow, 9)) + _
'                                          Val(flxBankPay.TextMatrix(iRow, 11)), "0.00")
'      flxBankPay.TextMatrix(iRow, 13) = adoRst!MY_ID
'      flxBankPay.TextMatrix(iRow, 14) = IIf(IsNull(adoRst!UNIT_ID), "", adoRst!UNIT_ID)
'      flxBankPay.TextMatrix(iRow, 15) = IIf(IsNull(adoRst!propertyID), "", adoRst!propertyID)
'      flxBankPay.TextMatrix(iRow, 16) = adoRst!DEPT_ID
'      flxBankPay.TextMatrix(iRow, 17) = IIf(IsNull(adoRst!ReconNow), "N", "Y")
'      flxBankPay.TextMatrix(iRow, 18) = adoRst!ClientName
'
'      adoRst.MoveNext
'      If Not adoRst.EOF Then flxBankPay.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'   adoRst.Close
'   Set adoRst = Nothing
End Sub

'TRAN_ID, BANK_AC, TRANS, TRAN_DATE, UNIT_ID, NOMINAL_CODE, PROJ_REF, FundName,
'description, NET_AMOUNT, TAX_CODE, VAT, , MY_ID,Client, PropID, FundID
Private Sub cmdEdit_Click()
'   If flxBankPay.TextMatrix(flxBankPay.row, 17) = "Y" Then
'      MsgBox "The transaction has been bank reconciled."
'      Exit Sub
'   End If
'   Load frmBankTranEdit
'   frmBankTranEdit.Caption = "Edit Bank Transaction"
'   frmBankTranEdit.FrmBankTranEdit_CALLING_FROM = Me.Name
'
'   With frmBankTranEdit
'      .szTransID = flxBankPay.TextMatrix(flxBankPay.row, 13)
'      .txtClientList.Tag = flxBankPay.TextMatrix(flxBankPay.row, 4)
'      .txtClientList.text = flxBankPay.TextMatrix(flxBankPay.row, 18)
'      .cboBC.Value = flxBankPay.TextMatrix(flxBankPay.row, 1)
'      .cboProperty.Value = flxBankPay.TextMatrix(flxBankPay.row, 15)
'      .cboUnit.Value = flxBankPay.TextMatrix(flxBankPay.row, 14)
'      .txtDetails.text = flxBankPay.TextMatrix(flxBankPay.row, 8)
'      .txtReference.text = flxBankPay.TextMatrix(flxBankPay.row, 6)
'      .cboNC.Value = flxBankPay.TextMatrix(flxBankPay.row, 5)
'      .cboFund.Value = flxBankPay.TextMatrix(flxBankPay.row, 16)
'      .txtDate.text = flxBankPay.TextMatrix(flxBankPay.row, 3)
'      .txtNet.text = Format(flxBankPay.TextMatrix(flxBankPay.row, 9), "0.00")
'      .cboVat.Value = flxBankPay.TextMatrix(flxBankPay.row, 10)
'      .txtTotal.text = Format(Val(flxBankPay.TextMatrix(flxBankPay.row, 9)) + _
'                       Val(flxBankPay.TextMatrix(flxBankPay.row, 11)), "0.00")
'   End With
'
'   frmBankTranEdit.Show
'   Me.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)

   If bFrmCashBookOpen Then
      frmCashbook.Show
   End If
   If BANK_PAYMENT_HISTORY_LOADED And sCallingForm = 1 Then
      frmDemands3.ZOrder 0
      frmDemands3.Show
      Me.Hide
   End If
   If BANK_PAYMENT_HISTORY_LOADED And sCallingForm = 2 Then
      frmBankTransactions.ZOrder 0
      frmBankTransactions.Show
      Me.Hide
   End If
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
End Sub
