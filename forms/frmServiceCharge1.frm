VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmServiceCharge1 
   Caption         =   "Service Charge Budget"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12570
   Icon            =   "frmServiceCharge1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12570
   Begin VB.CommandButton Command1 
      Caption         =   "&Import"
      Height          =   345
      Left            =   6480
      TabIndex        =   23
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdNew 
      Caption         =   "&New"
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdEdit 
      Caption         =   "&Edit"
      Height          =   345
      Left            =   1450
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   2825
      TabIndex        =   6
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSCBdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   4080
      TabIndex        =   7
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   8760
      TabIndex        =   8
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdClose 
      Caption         =   "Cl&ose"
      Height          =   345
      Left            =   10140
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Index           =   1
      Left            =   75
      TabIndex        =   10
      Top             =   75
      Width           =   11160
      Begin VB.TextBox txtBudgetId 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   5880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtMatrixRow 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   5880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDetails 
         Caption         =   ">>"
         Height          =   315
         Left            =   6400
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtBudget 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   295
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   380
         Width           =   2200
      End
      Begin VB.TextBox txtPpsf 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txtRentChargesIDEdit 
         Height          =   285
         Left            =   12720
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtTotalArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   360
         Width           =   1905
      End
      Begin VB.TextBox txtSCTotalArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   5895
         Width           =   1935
      End
      Begin VB.TextBox txtSCBudgetTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   11
         Top             =   5895
         Width           =   2715
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCBudgetDetails 
         Height          =   5145
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   9075
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   0
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
         _Band(0).Cols   =   6
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Area"
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price/SqFoot"
         Height          =   195
         Index           =   4
         Left            =   8880
         TabIndex        =   18
         Top             =   120
         Width           =   945
      End
      Begin MSForms.ComboBox cboSCFund 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
         VariousPropertyBits=   1755334683
         DisplayStyle    =   7
         Size            =   "7223;556"
         TextColumn      =   2
         ColumnCount     =   6
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705;70555"
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         Caption         =   "Total Budget"
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   17
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         Caption         =   "Fund"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   12
         Left            =   3600
         TabIndex        =   15
         Top             =   5895
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmServiceCharge1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn    As New ADODB.Connection
Dim szSQL   As String
Dim flgEdit As Integer
Dim flgNew  As Integer
Dim nextRow As Integer

Dim flgChange              As Integer
Dim detailsMatrix(59, 59)  As clsSCDtl
Dim matrixLength           As Integer
Dim bufferMatrix()         As clsSCDtl

Public Function fillBufferMatrix(ByVal bDetails As clsSCDtl, ByVal num As Integer)
   ReDim Preserve bufferMatrix(num) As clsSCDtl

   Set bufferMatrix(num - 1) = New clsSCDtl

   bufferMatrix(num - 1).setBudgetAmount bDetails.getBudgetAmount
   bufferMatrix(num - 1).setBudgetDetailId bDetails.getBudgetDetailID
   bufferMatrix(num - 1).setBudgetId bDetails.getBudgetId
   bufferMatrix(num - 1).setFlgDel bDetails.getFlgDel
   bufferMatrix(num - 1).setNCode bDetails.getNCode
   bufferMatrix(num - 1).setNName bDetails.getNName
End Function

Private Sub cboSCFund_Change()
   If Not txtBudget.Locked Then
      txtBudget.SetFocus
   Else
      If Not txtTotalArea.Locked Then
         frmServiceCharge1.Show
         txtTotalArea.SetFocus
      End If
   End If
End Sub

Private Sub cmdDetails_Click()
   Load frmServiceChargeDetails
   frmServiceChargeDetails.Show
   If flgNew = 1 And Val(Trim(txtBudget.text)) = 0 Then frmServiceChargeDetails.initialiseNew
   Me.Enabled = False
End Sub

Public Function initialiseMatrix() As Integer
   Dim i As Integer, j As Integer

   For i = 0 To 59
      For j = 0 To 59
         Set detailsMatrix(i, j) = New clsSCDtl
      Next j
   Next i
End Function

Public Function getDetailsFromMatrix(ByVal row As Integer, ByVal col As Integer) As clsSCDtl
   Set getDetailsFromMatrix = detailsMatrix(row, col)
End Function

Private Sub cmdSCBDCancel_Click()
   If cmdSCBdEdit.Caption = "&Update" Then
      If MsgBox("Do you want to cancel the operation?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
         rollBackServiceChargeDetails
         ReDim bufferMatrix(0) As clsSCDtl
         ControlsModeRentBudgetDetails DefaultMode
         Exit Sub
      End If
   End If

   If MsgBox("Are you sure to cancel the changes made since the last Save?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
         flxSCBudgetDetails.Clear
         initialiseGrid
         ControlsModeRentBudgetDetails DefaultMode
         SCSumTotal
   End If
End Sub

Private Sub rollBackServiceChargeDetails()
   Dim i As Integer

   For i = 0 To 59
      If detailsMatrix(txtMatrixRow.text, i).getBudgetId <> "" Then
         Set detailsMatrix(txtMatrixRow.text, i) = New clsSCDtl
      Else
         Exit For
      End If
   Next i

   If Not IsEmpty(bufferMatrix) Then
      For i = 0 To UBound(bufferMatrix) - 1
         Set detailsMatrix(txtMatrixRow.text, i) = bufferMatrix(i)
      Next i
   End If
End Sub

Private Sub cmdSCBdClose_Click()
   If flgChange = 1 Then
      If MsgBox("Do you want to save your changes before closing?", vbQuestion + vbYesNo, "New Budget") = vbYes Then
         cmdSCBDSave_Click
      End If
   End If

   Me.Hide
   Unload Me
End Sub

Private Sub cmdSCBdDelete_Click()
   If MsgBox("Do you want to delete the selected Rent Charge Budget Detail?", vbQuestion + vbYesNo, "Saving") = vbNo Then
         Exit Sub
   End If
   If Trim(flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 0)) <> "" Then
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 7) = "X"
      flxSCBudgetDetails.RowHeight(flxSCBudgetDetails.row) = 0
   Else
      flxSCBudgetDetails.RemoveItem (flxSCBudgetDetails.row)
   End If
   flgChange = 1
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Sub cmdSCBDEdit_Click()
   Dim col As Integer

   If cmdSCBdEdit.Caption = "&Edit" Then
      flgEdit = 1
      ControlsModeRentBudgetDetails EditMode

      If hasNoLines Then
         cmdDetails.Enabled = False
      Else
         cmdDetails.Enabled = True
         txtBudget.Locked = True
      End If

      cboSCFund.SetFocus

      For col = 0 To 59
         If getDetailsFromMatrix(txtMatrixRow.text, col).getBudgetDetailID = "" Then
            Exit For
         Else
            fillBufferMatrix getDetailsFromMatrix(txtMatrixRow.text, col), col + 1
         End If
      Next col
   Else
      If cboSCFund.Value = "" Then
         ShowMsgInTaskBar "Please select the fund.", , "N"
         cboSCFund.SetFocus
         Exit Sub
      End If
      If txtBudget.text = "" Then
         ShowMsgInTaskBar "Please enter the total budget.", , "N"
         txtBudget.SetFocus
         Exit Sub
      End If
      If txtTotalArea.text = "" Then
         ShowMsgInTaskBar "Please enter the total area.", , "N"
         txtTotalArea.SetFocus
         Exit Sub
      End If

      updateGrid

      ReDim bufferMatrix(0) As clsSCDtl
      flgChange = 1
      If flgNew = 1 Then
         nextRow = nextRow + 1
         flgNew = 0
      End If

      SCSumTotal
      ControlsModeRentBudgetDetails DefaultMode
   End If
End Sub

Private Sub cmdSCBdNew_Click()
   

   flgChange = 1
   txtMatrixRow.text = nextRow
   txtBudgetId = UniqueID()
   flgNew = 1
   ControlsModeRentBudgetDetails NewEntryMode

   If MsgBox("Do you wish to analyse the budget?", vbQuestion + vbYesNo, "Analyse Budget") = vbYes Then
      txtBudget.Locked = True
      cmdDetails_Click
      
   Else
      cmdDetails.Enabled = False
      txtBudget.Locked = False
      cboSCFund.DropDown
      cboSCFund.SetFocus
   End If
End Sub

Private Sub cmdSCBDSave_Click()
   Dim i          As Integer
   Dim iRow       As Integer
   Dim iRowChild  As Integer
   Dim Rst        As New ADODB.Recordset

   adoConn.Open getConnectionString

   i = 0
   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.TextMatrix(iRow, 7) <> "X" Then
         SaveUpdateSC iRow
         txtBudgetId.text = flxSCBudgetDetails.TextMatrix(iRow, 0)
         saveSCDetails i
      Else
         deleteSC flxSCBudgetDetails.TextMatrix(iRow, 0)
      End If
   i = i + 1
   Next iRow

   Update_SC_Lease

'  Update Service Charge in LServiceCharge table where 'ChargingMethod' were 'Percentage -> 2'

'   szSQL = "UPDATE LServiceCharges AS LSC, GlobalSC AS GSC, Frequencies AS F, " & _
'              "LeaseDetails AS L, Units AS U " & _
'            "SET LSC.SCTotal = GSC.TotalBudget * (LSC.CMFigure / 100), " & _
'              "LSC.SCAmount = (GSC.TotalBudget * (LSC.CMFigure / 100)) / F.PartOfYear " & _
'            "WHERE LSC.ChargingMethod = 2 AND CByte(LSC.ServiceChargeDept) = GSC.Fund AND " & _
'              "L.LeaseID = LSC.LeaseID AND L.UnitNumber = U.UnitNumber AND " & _
'              "U.PropertyID = '" & frmGlobal1.cboProperty.BoundText & "' AND " & _
'              "LSC.SCFrequency = F.ID AND GSC.PropertyID = '" & frmGlobal1.cboProperty.BoundText & "';"
'   szSQL = "UPDATE LServiceCharges AS LSC, GlobalSC AS GSC, Frequencies AS F, " & _
'              "LeaseDetails AS L, Units AS U " & _
'           "SET LSC.SCTotal = GSC.TotalBudget * (LSC.CMFigure / 100), " & _
'              "LSC.SCAmount = (GSC.TotalBudget * (LSC.CMFigure / 100)) / F.PartOfYear " & _
'           "WHERE LSC.ChargingMethod = 2 AND CByte(LSC.ServiceChargeDept) = GSC.Fund AND " & _
'              "L.LeaseID = LSC.LeaseID AND L.UnitNumber = U.UnitNumber AND " & _
'              "U.PropertyID = GSC.PropertyID AND LSC.SCFrequency = F.ID;"
''Debug.Print szSQL
'   adoConn.Execute szSQL
'
''  Update Service Charge in LServiceCharge table where 'ChargingMethod' were 'Global -> 4'
'   szSQL = "UPDATE LServiceCharges AS LSC, GlobalSC AS GSC, Frequencies AS F, LeaseDetails AS L, Units AS U " & _
'           "SET LSC.SCTotal = GSC.PPSF * U.TotalArea, " & _
'              "LSC.SCAmount = GSC.PPSF * U.TotalArea / F.PartOfYear, " & _
'              "LSC.CMFigure = GSC.PPSF * U.TotalArea " & _
'           "WHERE LSC.ChargingMethod = 4 AND CByte(LSC.ServiceChargeDept) = GSC.Fund AND " & _
'              "LSC.SCFrequency = F.ID AND LSC.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber AND " & _
'              "U.PropertyID = GSC.PropertyID"
''Debug.Print szSQL
'   adoConn.Execute szSQL
'
   loadMatrix adoConn

   frmGlobal1.lblSCTotal.Caption = Format(TotalSCProperty(frmGlobal1.cboProperty.BoundText), "£0.00")

   Set Rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
   flgEdit = 0
   flgNew = 0
   flgChange = 0

   ControlsModeRentBudgetDetails SavedMode
End Sub

Private Sub Update_SC_Lease()
   Dim Rst     As New ADODB.Recordset
   Dim adoRst  As New ADODB.Recordset

   szSQL = "SELECT * FROM LServiceCharges WHERE ChargingMethod = 4 OR ChargingMethod = 2;"

   With Rst
      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      While Not .EOF
         If .Fields.Item("ChargingMethod").Value = 2 Then
            szSQL = "SELECT GSC.TotalBudget, F.PartOfYear " & _
                    "FROM   GlobalSC AS GSC, LeaseDetails AS L, Units AS U, " & _
                        "Frequencies AS F, LServiceCharges AS LSC " & _
                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
                        "L.UnitNumber = U.UnitNumber AND " & _
                        "L.LeaseID = LSC.LeaseID AND " & _
                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
                        "LSC.SCFrequency = F.ID AND " & _
                        "U.PropertyID = GSC.PropertyID;"
'Debug.Print szSQL
            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            
            If Not adoRst.EOF Then
               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("TotalBudget").Value * _
                                                .Fields.Item("CMFigure").Value / 100
               .Fields.Item("SCAmount").Value = (adoRst.Fields.Item("TotalBudget").Value * _
                                                (.Fields.Item("CMFigure").Value / 100)) / _
                                                adoRst.Fields.Item("PartOfYear").Value
            Else
               .Fields.Item("SCTotal").Value = 0
               .Fields.Item("SCAmount").Value = 0
            End If
            adoRst.Close
         Else                             '.ChargingMethod = 4
            szSQL = "SELECT GSC.PPSF, U.TotalArea, F.PartOfYear " & _
                    "FROM GlobalSC AS GSC, LeaseDetails AS L, Units AS U " & _
                        "Frequencies AS F, LServiceCharges AS LSC " & _
                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
                        "L.UnitNumber = U.UnitNumber AND " & _
                        "L.LeaseID = LSC.LeaseID AND " & _
                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
                        "LSC.SCFrequency = F.ID AND " & _
                        "U.PropertyID = GSC.PropertyID;"
'Debug.Print szSQL
            If Not adoRst.EOF Then
               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("PPSF").Value * _
                                               adoRst.Fields.Item("TotalArea").Value
               .Fields.Item("SCAmount").Value = adoRst.Fields.Item("PPSF").Value * _
                                                adoRst.Fields.Item("TotalArea").Value / _
                                                adoRst.Fields.Item("PartOfYear").Value
               .Fields.Item("CMFigure").Value = adoRst.Fields.Item("PPSF").Value * _
                                                adoRst.Fields.Item("TotalArea").Value

            Else
               .Fields.Item("SCTotal").Value = 0
               .Fields.Item("SCAmount").Value = 0
               .Fields.Item("CMFigure").Value = 0
            End If
            adoRst.Close
         End If
         .Update
         .MoveNext
      Wend
      .Close
   End With

   Set adoRst = Nothing
   Set Rst = Nothing
End Sub

Private Function saveSCDetails(ByVal row As Integer)
   Dim col As Integer

   For col = 0 To 59
      If frmServiceCharge1.getDetailsFromMatrix(row, col).getBudgetDetailID = "" Then
          Exit For
      Else
         If frmServiceCharge1.getDetailsFromMatrix(row, col).getFlgDel = "X" Then
            deleteSCD frmServiceCharge1.getDetailsFromMatrix(row, col)
         Else
            SaveUpdateSCD frmServiceCharge1.getDetailsFromMatrix(row, col)
         End If
      End If
   Next col
End Function

Private Function SaveUpdateSCD(ByVal budgetDetails As clsSCDtl)
   Dim Rst     As New ADODB.Recordset

   szSQL = "DELETE * " & _
            "FROM GlobalSCDtls " & _
            "WHERE BudgetDtlID = '" & budgetDetails.getBudgetDetailID & "' "
   Rst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    
  
   szSQL = "SELECT * " & _
            "FROM GlobalSCDtls;"
   Rst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   
   Rst.AddNew
   Rst!BudgetDtlID = budgetDetails.getBudgetDetailID
   Rst!budgetId = budgetDetails.getBudgetId
   Rst!NN = budgetDetails.getNName
   Rst!NC = budgetDetails.getNCode
   Rst!BudgetAmt = budgetDetails.getBudgetAmount
   
   Rst.Update
   Rst.Close
   Set Rst = Nothing
End Function

Private Function deleteSCD(ByVal budgetDetails As clsSCDtl)
   szSQL = "DELETE * " & _
            "FROM GlobalSCDtls " & _
            "WHERE BudgetDtlID = '" & budgetDetails.getBudgetDetailID & "' "
   adoConn.Execute szSQL
End Function

Private Function SaveUpdateSC(ByVal Index As Integer)
   Dim Rst     As New ADODB.Recordset
   
   szSQL = "SELECT * " & _
            "FROM GlobalSC where budgetId='" & flxSCBudgetDetails.TextMatrix(Index, 0) & "';"
   Rst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   If Rst.RecordCount < 1 Then
      Rst.AddNew
      Rst!budgetId = flxSCBudgetDetails.TextMatrix(Index, 0)
   End If
   
   Rst!propertyID = frmGlobal1.cboProperty.BoundText
   Rst!Fund = flxSCBudgetDetails.TextMatrix(Index, 2)
   Rst!TotalBudget = flxSCBudgetDetails.TextMatrix(Index, 4)
   Rst!SCArea = flxSCBudgetDetails.TextMatrix(Index, 5)
   Rst!ppsf = flxSCBudgetDetails.TextMatrix(Index, 6)

   Rst.Update
   Rst.Close
   Set Rst = Nothing
End Function

Private Function deleteSC(ByVal bId As String)
   szSQL = "DELETE * " & _
            "FROM GlobalSC " & _
            "WHERE BudgetID = '" & bId & "';"
   adoConn.Execute szSQL

'   szSQL = "DELETE GlobalSCDtls.* " & _
'           "FROM   GlobalSCDtls " & _
'           "WHERE  GlobalSCDtls.BudgetID NOT IN (" & _
'               "SELECT GlobalSC.BudgetID " & _
'               "FROM GlobalSC);"
'issue  381 SQL that is taking long time to complete 20170512 fixed by anol
   szSQL = "DELETE GlobalSCDtls.* From GlobalSCDtls LEFT JOIN GlobalSC ON GlobalSCDtls.BudgetID =GlobalSC.BudgetID WHERE  GlobalSC.BudgetID  IS NULL;"
   adoConn.Execute szSQL
End Function

Private Sub cmdSCClose_Click()
   initialiseGrid
   Me.Hide
End Sub


Private Sub ConfigureFlxBRMain()
   Dim szFlxHeader As String

   flxSCBudgetDetails.Rows = 1
   flxSCBudgetDetails.RowHeight(0) = 0
   flxSCBudgetDetails.Clear
   flxSCBudgetDetails.Cols = 9
   szFlxHeader$ = "BudgetID|PropertyID|<Fund|>FundName|>TotalBudget|>SCArea|>PPSF"
   flxSCBudgetDetails.FormatString = szFlxHeader$

   flxSCBudgetDetails.ColWidth(0) = 0
   flxSCBudgetDetails.ColWidth(1) = 0
   flxSCBudgetDetails.ColWidth(2) = 0
   flxSCBudgetDetails.ColWidth(3) = lblRentCharges(2).Left - lblRentCharges(0).Left
   flxSCBudgetDetails.ColWidth(4) = lblRentCharges(3).Left - lblRentCharges(2).Left
   flxSCBudgetDetails.ColWidth(5) = lblRentCharges(4).Left - lblRentCharges(3).Left
   flxSCBudgetDetails.ColWidth(6) = flxSCBudgetDetails.Width - lblRentCharges(4).Left - 300
   flxSCBudgetDetails.ColWidth(7) = 0
   flxSCBudgetDetails.ColWidth(8) = 0
End Sub

Private Sub LoadFlxSCMain(Conn1 As ADODB.Connection)
   Dim i       As Integer
   Dim Rst     As New ADODB.Recordset

   szSQL = "SELECT g.*,f.FundName " & _
            "FROM GlobalSC g, Fund f " & _
            "WHERE CInt(g.Fund)=f.FundId AND PropertyID = '" & frmGlobal1.cboProperty.BoundText & "';"
   Rst.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   i = 1
   If Not Rst.EOF Then
      While Not Rst.EOF
         flxSCBudgetDetails.AddItem ""
         flxSCBudgetDetails.TextMatrix(i, 0) = Rst!budgetId
         flxSCBudgetDetails.TextMatrix(i, 1) = Rst!propertyID
         flxSCBudgetDetails.TextMatrix(i, 2) = Rst!Fund
         flxSCBudgetDetails.TextMatrix(i, 3) = Rst!FundName
         flxSCBudgetDetails.TextMatrix(i, 4) = Format(Rst!TotalBudget, "0.00")
         flxSCBudgetDetails.TextMatrix(i, 5) = Rst!SCArea
         flxSCBudgetDetails.TextMatrix(i, 6) = Format(Rst!ppsf, "0.00")
         flxSCBudgetDetails.TextMatrix(i, 8) = i - 1
         i = i + 1
         Rst.MoveNext
      Wend
   End If
   nextRow = i - 1
   initialiseMatrix
   flxSCBudgetDetails.row = 0
   flxSCBudgetDetails.col = 0

   Rst.Close
   Set Rst = Nothing
End Sub

Private Sub loadMatrix(Conn1 As ADODB.Connection)
   Dim i As Integer

   For i = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.TextMatrix(i, 0) <> "" Then
         PopulateMatrix flxSCBudgetDetails.TextMatrix(i, 0), flxSCBudgetDetails.TextMatrix(i, 8), Conn1
      End If
   Next i
End Sub

Private Sub PopulateMatrix(bId As String, row As Integer, Conn1 As ADODB.Connection)
   Dim i    As Integer
   Dim Rst  As New ADODB.Recordset

   szSQL = "SELECT g.* " & _
            "FROM   GlobalSCDtls AS g " & _
            "WHERE  g.BudgetID = '" & bId & "';"
   Rst.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   i = 0
   If Not Rst.EOF Then
      While Not Rst.EOF
         getDetailsFromMatrix(row, i).setBudgetDetailId Rst!BudgetDtlID
         getDetailsFromMatrix(row, i).setBudgetId Rst!budgetId
         getDetailsFromMatrix(row, i).setNCode Rst!NC
         getDetailsFromMatrix(row, i).setNName Rst!NN
         getDetailsFromMatrix(row, i).setBudgetAmount Format(Rst!BudgetAmt, "0.00")
         i = i + 1
         Rst.MoveNext
      Wend
   End If

   Rst.Close
   Set Rst = Nothing
End Sub

Private Sub SCSumTotal()
   Dim iRow As Integer

   txtSCBudgetTotal.text = "0.00"
   txtSCTotalArea.text = "0"
   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      txtSCBudgetTotal.text = Format(Val(txtSCBudgetTotal.text) + Val(flxSCBudgetDetails.TextMatrix(iRow, 4)), "0.00")
      txtSCTotalArea.text = Val(txtSCTotalArea.text) + Val(flxSCBudgetDetails.TextMatrix(iRow, 5))
   Next iRow
End Sub

Private Sub flxSCBudgetDetails_RowColChange()
   If flxSCBudgetDetails.TextMatrix(1, 0) = "" Then Exit Sub
   If flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 7) = "X" Then Exit Sub

   On Error Resume Next

   cboSCFund.Value = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 2)
   txtBudget.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 4)
   txtTotalArea.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 5)
   txtPpsf.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 6)
   txtMatrixRow.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 8)
   txtBudgetId.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 0)

   ControlsModeRentBudgetDetails GridRowOnSelection

   cmdSCBdEdit.SetFocus
End Sub
Private Function hasNoLines() As Boolean
   Dim i As Integer

   hasNoLines = True
   i = 0
   While i < 60
      If detailsMatrix(flxSCBudgetDetails.row - 1, i).getBudgetId <> "" Then
         If detailsMatrix(flxSCBudgetDetails.row - 1, i).getFlgDel <> "X" Then
            hasNoLines = False
            Exit Function
         End If
      Else
         Exit Function
      End If
      i = i + 1
   Wend
End Function


Private Sub ControlsModeRentBudgetDetails(ByVal mode As ComponentMode)
   On Error Resume Next

   Select Case mode
      Case ComponentMode.DefaultMode
         cboSCFund.text = ""
         cboSCFund.Enabled = False
         txtBudget.text = ""
         txtBudgetId.text = ""
         txtMatrixRow.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True

         cmdDetails.Enabled = False
         cmdSCBdNew.Enabled = True
         cmdSCBdEdit.Enabled = False
         cmdSCBdEdit.Caption = "&Edit"
         cmdSCBdEdit.ToolTipText = "Edit selected budget"
         If flgChange = 1 Then
            cmdSCBdSave.Enabled = True
            cmdSCBdCancel.Enabled = True
         Else
            cmdSCBdSave.Enabled = False
            cmdSCBdCancel.Enabled = False
         End If
         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = True
         flxSCBudgetDetails.row = 0
         flxSCBudgetDetails.col = 0
         cmdSCBdNew.SetFocus

      Case ComponentMode.SavedMode
         cboSCFund.text = ""
         cboSCFund.Enabled = False
         txtBudget.text = ""
         txtBudgetId.text = ""
         txtMatrixRow.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True

         cmdDetails.Enabled = False
         cmdSCBdNew.Enabled = True
         cmdSCBdEdit.Enabled = False
         cmdSCBdEdit.Caption = "&Edit"
         cmdSCBdEdit.ToolTipText = "Edit selected budget"
         If flgChange = 1 Then
            cmdSCBdSave.Enabled = True
            cmdSCBdCancel.Enabled = True
         Else
            cmdSCBdSave.Enabled = False
            cmdSCBdCancel.Enabled = False
         End If

         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = True
         flxSCBudgetDetails.row = 0
         flxSCBudgetDetails.col = 0
         cmdSCBdClose.SetFocus

      Case ComponentMode.EditMode

         cboSCFund.Enabled = True
         txtTotalArea.Locked = False
         txtBudget.Locked = False
         cmdDetails.Enabled = True
         cmdSCBdNew.Enabled = False
         cmdSCBdEdit.Caption = "&Update"
         cmdSCBdEdit.ToolTipText = "Update selected budget"
         cmdSCBdEdit.Enabled = True
         cmdSCBdSave.Enabled = False
         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = False

      Case ComponentMode.NewEntryMode
         cboSCFund.text = ""
         cboSCFund.Enabled = True
         txtBudget.text = ""
         txtTotalArea.text = ""
         txtTotalArea.Locked = False
         txtPpsf.text = ""

         cmdDetails.Enabled = True
         cmdSCBdNew.Enabled = False
         cmdSCBdEdit.Caption = "&Update"
         cmdSCBdEdit.Enabled = True
         cmdSCBdSave.Enabled = False
         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = False
         flxSCBudgetDetails.Enabled = False

      Case ComponentMode.GridRowOnSelection
         txtBudget.Locked = True
         cmdSCBdNew.Enabled = True
         cmdSCBdEdit.Enabled = True

        If flgChange = 1 Then
            cmdSCBdSave.Enabled = True
         Else
            cmdSCBdSave.Enabled = False
         End If
         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = True
   End Select
End Sub

Private Sub Form_Load()
   Me.Width = 11460
   Me.Height = 7800
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Refresh
   Me.BackColor = MODULEBACKCOLOR
   Frame1(1).BackColor = MODULEBACKCOLOR

   initialiseGrid
   SCSumTotal

   Call WheelHook(Me.hWnd)
End Sub

Private Sub initialiseGrid()
   Call ConfigureFlxBRMain

   flgEdit = 0
   flgChange = 0
   flgNew = 0

   adoConn.Open getConnectionString

   LoadFlxSCMain adoConn
   loadMatrix adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call LoadFund
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Function TotalSCProperty(szPropertyID As String) As Double
   Dim Rst2 As New ADODB.Recordset

   szSQL = "SELECT GlobalSC.PropertyID, SUM(GlobalSC.TotalBudget) AS TOTALRENT " & _
            "From GlobalSC " & _
            "WHERE GlobalSC.PropertyID = '" & szPropertyID & "' " & _
            "GROUP BY GlobalSC.PropertyID;"
'Debug.Print szSQL
   Rst2.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not Rst2.EOF Then
      TotalSCProperty = CDbl(Rst2!TOTALRENT)
   Else
      TotalSCProperty = 0
   End If

   Rst2.Close
   Set Rst2 = Nothing
End Function

Private Sub LoadFund()
   ' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund " & _
           "WHERE CategoryCode = 2;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
   Else
      ReDim Data(2, adoRst.RecordCount) As String
    
      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = Trim(adoRst.Fields.Item("FundID").Value)
         Data(1, rRow) = Trim(adoRst.Fields.Item("FundName").Value)
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
    
      cboSCFund.Clear
      cboSCFund.Column() = Data()
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading fund.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   frmGlobal1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtBudget_Change()
   If Trim(txtBudget.text) <> "" And Val(Trim(txtBudget.text)) > 0 Then
    txtBudget.Locked = False
    computePpsf
   End If
End Sub

Private Sub txtBudget_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtBudget, KeyAscii
End Sub

Private Sub txtBudget_LostFocus()
   computePpsf
'   If flgEdit = 1 Or flgNew = 1 Then txtBudget.BackColor = &H80000005
End Sub

Private Sub computePpsf()
On Error GoTo ErrHandler
   If Trim(txtBudget.text) <> "" And Trim(txtTotalArea.text) <> "" Then
      Dim ppsf As Double
      
      ppsf = CDbl(txtBudget.text) / CDbl(Trim(txtTotalArea.text))
      txtPpsf.text = FormatNumber(CDbl(ppsf), 2, , , vbDefault)
   End If
   Exit Sub
ErrHandler:
   ShowMsgInTaskBar "Please ensure that the values for Total Budget and Total Area are valid before continuing", , "N"
End Sub

Private Sub txtTotalArea_LostFocus()
   computePpsf
End Sub

Private Sub updateGrid()
    If flgEdit = 0 Then
         flxSCBudgetDetails.AddItem ""
         If flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 0) <> "" Then flxSCBudgetDetails.AddItem ""
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 0) = txtBudgetId.text
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 1) = frmGlobal1.cboProperty.BoundText
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 2) = CInt(cboSCFund.Value)
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 3) = cboSCFund.text
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 5) = txtTotalArea.text
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 6) = txtPpsf.text
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 8) = nextRow
   Else
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = frmGlobal1.cboProperty.BoundText
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 2) = CInt(cboSCFund.Value)
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 3) = cboSCFund.text
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 5) = txtTotalArea.text
         flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 6) = txtPpsf.text
   End If
      flgEdit = 0
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
