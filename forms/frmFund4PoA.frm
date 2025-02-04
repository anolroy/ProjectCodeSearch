VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFund4PoA 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9075
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFund4PoA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFund 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   960
      ScaleHeight     =   1575
      ScaleWidth      =   4095
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   4095
      Begin MSForms.ComboBox cboFund 
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   2925
         VariousPropertyBits=   1753237523
         BackColor       =   15066613
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5159;556"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705"
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   3720
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3720
      Width           =   1185
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPoAFundSetup 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   8
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      Height          =   195
      Index           =   7
      Left            =   7800
      TabIndex        =   8
      Top             =   0
      Width           =   360
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      Height          =   195
      Index           =   6
      Left            =   6960
      TabIndex        =   7
      ToolTipText     =   "Receipt Date"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Index           =   5
      Left            =   6000
      TabIndex        =   6
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   600
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   0
      Width           =   285
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Receipt Date"
      Top             =   0
      Width           =   345
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   195
      Index           =   3
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblGridLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   195
      Index           =   4
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "frmFund4PoA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bUD As Boolean

Private Sub cboFund_Click()
   If bUD Then
      bUD = False
      Exit Sub
   End If

   If picFund.Visible Then _
      flxPoAFundSetup.TextMatrix(flxPoAFundSetup.row, 7) = cboFund.Column(0)
   picFund.Visible = False
   flxPoAFundSetup.SetFocus
End Sub

Private Sub cboFund_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   If KeyCode = 38 Or KeyCode = 40 Then bUD = True
   If KeyCode = 13 Then Call cboFund_Click
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
   If MsgBox("Do you wish to save?", vbQuestion + vbYesNo, "Fund") = vbNo Then Exit Sub

   Dim iRow As Integer

   For iRow = 1 To flxPoAFundSetup.Rows - 1
      If flxPoAFundSetup.TextMatrix(iRow, 7) <> "" Then
         frmMMain.Conn1.Execute "UPDATE tlbPayment " & _
                  "SET FundID = " & flxPoAFundSetup.TextMatrix(iRow, 7) & " " & _
                  "WHERE TransactionID = " & flxPoAFundSetup.TextMatrix(iRow, 8) & ";"
      End If
   Next iRow

   MsgBox "Fund has been updated successfully.", vbInformation + vbOKOnly, "Fund"
   Unload Me
End Sub

Private Sub flxPoAFundSetup_DblClick()
   Dim i As Integer, iFlxSPayCol As Integer

   If flxPoAFundSetup.TextMatrix(flxPoAFundSetup.row, 0) = "" Then Exit Sub

   flxPoAFundSetup.col = 7

   picFund.Top = flxPoAFundSetup.CellTop + flxPoAFundSetup.Top
   picFund.Left = flxPoAFundSetup.CellLeft + flxPoAFundSetup.Left
   picFund.Width = flxPoAFundSetup.ColWidth(7)
   picFund.Height = flxPoAFundSetup.RowHeight(flxPoAFundSetup.row) - 15
   cboFund.ListIndex = FindComboIndex(cboFund, flxPoAFundSetup.TextMatrix(flxPoAFundSetup.row, 7), 0)
   picFund.Visible = True
   cboFund.SetFocus
End Sub

Private Sub flxPoAFundSetup_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then flxPoAFundSetup_DblClick
End Sub

Private Sub Form_Activate()
   Dim iRow As Integer, szSQL As String

   frmMMain.Rst1.MoveFirst
   iRow = 1

   With frmMMain.Rst1
      While Not .EOF
         flxPoAFundSetup.TextMatrix(iRow, 0) = .Fields.Item("SageAccountNumber").Value
         flxPoAFundSetup.TextMatrix(iRow, 1) = IIf(IsNull(.Fields.Item("UnitID").Value), "", .Fields.Item("UnitID").Value)
         flxPoAFundSetup.TextMatrix(iRow, 2) = .Fields.Item("PDate").Value
         flxPoAFundSetup.TextMatrix(iRow, 3) = IIf(IsNull(.Fields.Item("Details").Value), "", .Fields.Item("Details").Value)
         flxPoAFundSetup.TextMatrix(iRow, 4) = IIf(IsNull(.Fields.Item("ExtRef").Value), "", .Fields.Item("ExtRef").Value)
         flxPoAFundSetup.TextMatrix(iRow, 5) = .Fields.Item("Amount").Value
         flxPoAFundSetup.TextMatrix(iRow, 6) = .Fields.Item("BankCode").Value
         flxPoAFundSetup.TextMatrix(iRow, 8) = .Fields.Item("TransactionID").Value
         flxPoAFundSetup.RowHeight(iRow) = 285

         .MoveNext
         iRow = iRow + 1
         flxPoAFundSetup.AddItem ""
      Wend
      .Close
   End With

   Dim TotalCol As Integer
   Dim Data() As String
   Dim j As Integer

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund " & _
           "ORDER BY FundID;"

   frmMMain.Rst1.Open szSQL, frmMMain.Conn1, adOpenStatic, adLockReadOnly

   If frmMMain.Rst1.EOF Then
      Set frmMMain.Rst1 = Nothing
      Exit Sub
   End If

   iRow = frmMMain.Rst1.RecordCount
   TotalCol = frmMMain.Rst1.Fields.count - 1

   ReDim Data(TotalCol, iRow) As String

   For iRow = 0 To iRow - 1
      For j = 0 To TotalCol
          Data(j, iRow) = IIf(IsNull(frmMMain.Rst1.Fields(j).Value), "", frmMMain.Rst1.Fields(j).Value)
      Next j
      frmMMain.Rst1.MoveNext
      If frmMMain.Rst1.EOF Then Exit For
   Next iRow

   cboFund.Column() = Data()
   cboFund.ListIndex = 0
End Sub

Private Sub Form_Load()
   Me.Width = 10770
   Me.Height = 4245
   Me.BackColor = MODULEBACKCOLOR
   Me.Top = frmMMain.Top / 2 - (Me.Height / 2)
   Me.Left = frmMMain.Width / 2 - (Me.Width / 2)

   Dim szHeader As String, iCol As Integer

   szHeader$ = "<|<|<|<|<|>|<|<"
   flxPoAFundSetup.FormatString = szHeader$

   flxPoAFundSetup.Clear
   flxPoAFundSetup.Cols = 9
   flxPoAFundSetup.Rows = 2
   flxPoAFundSetup.RowHeight(0) = 0

   For iCol = 0 To flxPoAFundSetup.Cols - 3
      flxPoAFundSetup.ColWidth(iCol) = lblGridLabel(iCol + 1).Left - lblGridLabel(iCol).Left
   Next iCol

   flxPoAFundSetup.ColWidth(7) = flxPoAFundSetup.Width + flxPoAFundSetup.Left - lblGridLabel(7).Left - 320
   flxPoAFundSetup.ColWidth(8) = 0
End Sub

Private Sub picFund_Resize()
   cboFund.Width = picFund.Width
   cboFund.Height = picFund.Height
End Sub
