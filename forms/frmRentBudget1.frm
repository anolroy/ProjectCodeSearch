VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRentBudget1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Budget Details"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13125
   Icon            =   "frmRentBudget1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7095
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   11040
      Begin VB.CommandButton cmdRCBdClose 
         Caption         =   "Cl&ose"
         Height          =   345
         Left            =   9720
         TabIndex        =   9
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdRCBdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   8400
         TabIndex        =   8
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton cmdRCBdDelete 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   3960
         TabIndex        =   7
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton cmdRCBdSave 
         Caption         =   "&Save"
         Height          =   345
         Left            =   2760
         TabIndex        =   6
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdRCBdEdit 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   1440
         TabIndex        =   5
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton cmdRCBdNew 
         Caption         =   "&New"
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox txtRCBudgetTotal 
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
         TabIndex        =   19
         Top             =   6120
         Width           =   2715
      End
      Begin VB.TextBox txtRCTotalArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   6120
         Width           =   1935
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
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   2745
      End
      Begin VB.TextBox txtRentChargesIDEdit 
         Height          =   285
         Left            =   12720
         TabIndex        =   11
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRentBudgetDetails 
         Height          =   5265
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   9287
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   12
         Left            =   3600
         TabIndex        =   18
         Top             =   6240
         Width           =   390
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
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         Caption         =   "Total Budget"
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   15
         Top             =   120
         Width           =   915
      End
      Begin MSForms.ComboBox cboRCFund 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
         VariousPropertyBits=   1753237531
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
         BackStyle       =   0  'Transparent
         Caption         =   "Price/SqFoot"
         Height          =   195
         Index           =   4
         Left            =   8880
         TabIndex        =   14
         Top             =   120
         Width           =   945
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Area"
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRentBudget1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sModule As String

Dim Conn As New ADODB.Connection
Dim Rst As New ADODB.Recordset
Dim SQLStr As String
Dim flgEdit As Integer
Dim flgNew As Integer
Dim flgChange As Integer

Private Sub cmdRCBdCancel_Click()
   If cmdRCBdEdit.Caption = "&Update" Then
      If MsgBox("Do you want to cancel the operation?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
         ControlsModeRentBudgetDetails DefaultMode
         Exit Sub
      End If
   End If
   
   If MsgBox("Are you sure to cancel the changes made since the last Save?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      flxRentBudgetDetails.Clear
      initialiseGrid
      ControlsModeRentBudgetDetails DefaultMode
      RCSumTotal
   End If
End Sub

Private Sub cmdRCBdClose_Click()
   If flgChange = 1 Then
      If MsgBox("Do you want to save your changes before closing?", vbQuestion + vbYesNo, "New Budget") = vbYes Then
         cmdRCBdSave_Click
      End If
   End If

   Me.Hide
   Unload Me
End Sub

Private Sub cmdRCBdDelete_Click()
   If MsgBox("Do you want to delete the selected " & IIf(sModule = "RB", "Rent", "Insurance") & " Charge Budget Detail?", vbQuestion + vbYesNo, "Saving") = vbNo Then
      Exit Sub
   End If
   If Trim(flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 0)) <> "" Then
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 7) = "X"
      flxRentBudgetDetails.RowHeight(flxRentBudgetDetails.row) = 0
   Else
      flxRentBudgetDetails.RemoveItem (flxRentBudgetDetails.row)
   End If
   flgChange = 1
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Sub cmdRCBdEdit_Click()
   If cmdRCBdEdit.Caption = "&Edit" Then
      flgEdit = 1
      ControlsModeRentBudgetDetails EditMode
      cboRCFund.SetFocus
   Else
      If cboRCFund.Value = "" Then
         ShowMsgInTaskBar "Please select the fund."
         cboRCFund.SetFocus
         Exit Sub
      End If
      If txtBudget.text = "" Then
         ShowMsgInTaskBar "Please enter the total budget."
         txtBudget.SetFocus
         Exit Sub
      End If
      If txtTotalArea.text = "" Then
         ShowMsgInTaskBar "Please enter the total area."
         txtTotalArea.SetFocus
         Exit Sub
      End If

      updateGrid
      flgChange = 1
      RCSumTotal
      ControlsModeRentBudgetDetails DefaultMode
   End If
End Sub

Private Sub cmdRCBdNew_Click()
   flgNew = 1
   ControlsModeRentBudgetDetails NewEntryMode
   cboRCFund.SetFocus
   cboRCFund.DropDown
End Sub

Private Sub cmdRCBdSave_Click()
   Dim iRow As Integer, iRowChild As Integer

   Conn.Open getConnectionString

   For iRow = 1 To flxRentBudgetDetails.Rows - 1
      If flxRentBudgetDetails.TextMatrix(iRow, 7) <> "X" Then
         saveUpdateRC iRow
      Else
         deleteRC flxRentBudgetDetails.TextMatrix(iRow, 0)
      End If
   Next iRow

   If sModule = "IB" Then
      SQLStr = "UPDATE LInsuranceCharges AS I, GlobalInsurance AS G, Frequencies AS F, " & _
                  "LeaseDetails AS L, Units AS U " & _
               "SET I.TotalYearlyInsurance = FORMAT(G.Amount * I.ChargingFigure / 100, '0.00'), " & _
                  "I.InsuranceEachPeriod = FORMAT(G.Amount * I.ChargingFigure / 100 / F.PartOfYear, '0.00') " & _
               "WHERE I.ChargingType = 2 AND CByte(I.InsuranceDept) = G.FundType AND L.Status AND " & _
                  "I.InsuranceFrequency = F.ID AND I.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = G.PropertyID;"
'Debug.Print SQLStr
   Else
      SQLStr = "UPDATE LRentCharges AS R, GlobalRC AS G, Frequencies AS F, LeaseDetails AS L, Units AS U " & _
               "SET R.spare2 = CStr(FORMAT(G.PPSF * U.TotalArea, '0.00')), " & _
                  "R.BRTotal = FORMAT(G.PPSF * U.TotalArea, '0.00'), " & _
                  "R.BRAmount = FORMAT((G.PPSF * U.TotalArea) / F.PartOfYear, '0.00') " & _
               "WHERE R.spare1 = '4' AND R.RentChargeDept = G.Fund AND L.Status AND " & _
                  "R.BRFrequency = F.ID AND R.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = G.PropertyID;"
   End If
'Debug.Print SQLStr
   Conn.Execute SQLStr

   If sModule = "IB" Then
      frmGlobal1.Label2(7).Caption = Format(TotalRCProperty(frmGlobal1.cboProperty.BoundText), "£0.00")
   Else
      frmGlobal1.Label2(8).Caption = Format(TotalRCProperty(frmGlobal1.cboProperty.BoundText), "£0.00")
   End If

   Set Rst = Nothing
   Conn.Close
   Set Conn = Nothing

   flgEdit = 0
   flgNew = 0
   flgChange = 0
   ControlsModeRentBudgetDetails SavedMode
End Sub

Private Function saveUpdateRC(ByVal Index As Integer)
   If sModule = "RB" Then
      SQLStr = "DELETE * " & _
               "FROM GlobalRC " & _
               "WHERE BudgetID = '" & flxRentBudgetDetails.TextMatrix(Index, 0) & "' "
   Else
      SQLStr = "DELETE * " & _
               "FROM GlobalInsurance " & _
               "WHERE ID = " & IIf(flxRentBudgetDetails.TextMatrix(Index, 0) = "", 0, flxRentBudgetDetails.TextMatrix(Index, 0)) & " "
   End If
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If sModule = "RB" Then
      SQLStr = "SELECT * " & _
               "FROM GlobalRC;"
   Else
      SQLStr = "SELECT I.ID, I.PropertyID, I.FundType AS Fund, " & _
                  "I.Amount AS TotalBudget, I.PPSF, I.SCArea " & _
               "FROM GlobalInsurance AS I;"
   End If

   Rst.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic
   Rst.AddNew
   If sModule = "RB" Then
      Rst!budgetId = UniqueID()
   Else
      Rst!ID = SlNumber("IB", "GlobalInsurance", Conn)
   End If
   Rst!PropertyID = frmGlobal1.cboProperty.BoundText
   Rst!Fund = flxRentBudgetDetails.TextMatrix(Index, 2)
   Rst!TotalBudget = flxRentBudgetDetails.TextMatrix(Index, 4)
   Rst!SCArea = flxRentBudgetDetails.TextMatrix(Index, 5)
   Rst!ppsf = flxRentBudgetDetails.TextMatrix(Index, 6)
   Rst.Update
   Rst.Close
End Function

Private Function deleteRC(ByVal bId As String)
   If sModule = "RB" Then
      SQLStr = "DELETE * " & _
               "FROM GlobalRC " & _
               "WHERE BudgetID = '" & bId & "' "
   Else
      SQLStr = "DELETE * " & _
               "FROM GlobalInsurance " & _
               "WHERE ID = " & bId & ";"
   End If
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly
End Function
Private Sub cmdRCClose_Click()
 
 initialiseGrid
 Me.Hide
 

End Sub

Private Sub ConfigureFlxBRMain()
   Dim szFlxHeader As String

   flxRentBudgetDetails.Rows = 1
   flxRentBudgetDetails.RowHeight(0) = 0
   flxRentBudgetDetails.Clear
   flxRentBudgetDetails.Cols = 8
   szFlxHeader$ = "BudgetID|PropertyID|<Fund|>FundName|>TotalBudget|>SCArea|>PPSF"
   flxRentBudgetDetails.FormatString = szFlxHeader$

   flxRentBudgetDetails.ColWidth(0) = 0
   flxRentBudgetDetails.ColWidth(1) = 0
   flxRentBudgetDetails.ColWidth(2) = 0
   flxRentBudgetDetails.ColWidth(3) = lblRentCharges(2).Left - lblRentCharges(0).Left
   flxRentBudgetDetails.ColWidth(4) = lblRentCharges(3).Left - lblRentCharges(2).Left
   flxRentBudgetDetails.ColWidth(5) = lblRentCharges(4).Left - lblRentCharges(3).Left
   flxRentBudgetDetails.ColWidth(6) = flxRentBudgetDetails.Width - lblRentCharges(4).Left - 300
   flxRentBudgetDetails.ColWidth(7) = 0
End Sub

Private Sub LoadFlxRentBudgetDetails()
   Dim i As Integer

   If sModule = "RB" Then
      SQLStr = "SELECT g.*, f.FundName " & _
               "FROM GlobalRC AS g, Fund AS f " & _
               "WHERE CInt(g.Fund) = f.FundId AND " & _
                     "PropertyID = '" & frmGlobal1.cboProperty.BoundText & "';"
   Else
      SQLStr = "SELECT I.ID AS BudgetID, I.FundType AS Fund, I.Amount AS TotalBudget, " & _
                  "f.FundName, I.PropertyID, I.SCArea, I.PPSF " & _
               "FROM GlobalInsurance I, Fund AS f " & _
               "WHERE I.PropertyID = '" & frmGlobal1.cboProperty.BoundText & "' AND " & _
                     "I.FundType = f.FundID;"
   End If
'Debug.Print SQLStr
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   i = 1
   If Not Rst.EOF Then
      While Not Rst.EOF
         flxRentBudgetDetails.AddItem ""
         flxRentBudgetDetails.TextMatrix(i, 0) = Rst!budgetId
         flxRentBudgetDetails.TextMatrix(i, 1) = Rst!PropertyID
         flxRentBudgetDetails.TextMatrix(i, 2) = Rst!Fund
         flxRentBudgetDetails.TextMatrix(i, 3) = Rst!FundName
         flxRentBudgetDetails.TextMatrix(i, 4) = Format(Rst!TotalBudget, "0.00")
         flxRentBudgetDetails.TextMatrix(i, 5) = IIf(IsNull(Rst!SCArea), "", Rst!SCArea)
         flxRentBudgetDetails.TextMatrix(i, 6) = IIf(IsNull(Rst!ppsf), "", Rst!ppsf)

         i = i + 1
         Rst.MoveNext
      Wend
   End If
   flxRentBudgetDetails.row = 0
   flxRentBudgetDetails.col = 0

   Rst.Close
   Set Rst = Nothing
End Sub

Private Sub RCSumTotal()
   Dim iRow As Integer

   txtRCBudgetTotal.text = "0.00"
   txtRCTotalArea.text = "0"
   For iRow = 1 To flxRentBudgetDetails.Rows - 1
      txtRCBudgetTotal.text = Format(Val(txtRCBudgetTotal.text) + Val(flxRentBudgetDetails.TextMatrix(iRow, 4)), "0.00")
      txtRCTotalArea.text = Val(txtRCTotalArea.text) + Val(flxRentBudgetDetails.TextMatrix(iRow, 5))
   Next iRow
End Sub

Private Sub flxRentBudgetDetails_RowColChange()
   If flxRentBudgetDetails.TextMatrix(1, 2) = "" Then Exit Sub

   On Error Resume Next

   cboRCFund.Value = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 2)
   txtBudget.text = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 4)
   txtTotalArea.text = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 5)
   txtPpsf.text = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 6)
   

   ControlsModeRentBudgetDetails GridRowOnSelection
   cmdRCBdEdit.SetFocus
End Sub

Private Sub ControlsModeRentBudgetDetails(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         cboRCFund.ListIndex = -1
         cboRCFund.Locked = True
         txtBudget.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True

         cmdRCBdNew.Enabled = True
         cmdRCBdEdit.Enabled = False
         cmdRCBdEdit.Caption = "&Edit"
         If flgChange = 1 Then
            cmdRCBdSave.Enabled = True
            cmdRCBdCancel.Enabled = True
         Else
            cmdRCBdSave.Enabled = False
            cmdRCBdCancel.Enabled = False
         End If
         
         cmdRCBdDelete.Enabled = False

         flxRentBudgetDetails.Enabled = True
         flxRentBudgetDetails.row = 0
         flxRentBudgetDetails.col = 0
         Me.Show
         cmdRCBdNew.SetFocus
         
      Case ComponentMode.SavedMode
         cboRCFund.text = ""
         cboRCFund.Locked = True
         txtBudget.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True
         

         cmdRCBdNew.Enabled = True
         cmdRCBdEdit.Enabled = False
         cmdRCBdEdit.Caption = "&Edit"
         If flgChange = 1 Then
            cmdRCBdSave.Enabled = True
            cmdRCBdCancel.Enabled = True
         Else
            cmdRCBdSave.Enabled = False
            cmdRCBdCancel.Enabled = False
         End If
         
         cmdRCBdDelete.Enabled = False
         
         flxRentBudgetDetails.Enabled = True
         flxRentBudgetDetails.row = 0
         flxRentBudgetDetails.col = 0
         frmRentBudget1.Show
         
         cmdRCBdClose.SetFocus

      Case ComponentMode.EditMode
      
         cboRCFund.Locked = False
         txtBudget.Locked = False
         txtTotalArea.Locked = False
         
         cmdRCBdNew.Enabled = False
         cmdRCBdEdit.Caption = "&Update"
         cmdRCBdEdit.Enabled = True
         cmdRCBdSave.Enabled = False
         cmdRCBdCancel.Enabled = True
         cmdRCBdDelete.Enabled = False

         flxRentBudgetDetails.Enabled = False

      Case ComponentMode.NewEntryMode
         cboRCFund.ListIndex = -1
         cboRCFund.Locked = False
         txtBudget.text = ""
         txtBudget.Locked = False
         txtTotalArea.text = ""
         txtTotalArea.Locked = False
         txtPpsf.text = ""
         
         cmdRCBdNew.Enabled = False
         cmdRCBdEdit.Caption = "&Update"
         cmdRCBdEdit.Enabled = True
         cmdRCBdSave.Enabled = False
         cmdRCBdCancel.Enabled = True
         cmdRCBdDelete.Enabled = False

         flxRentBudgetDetails.Enabled = False

      Case ComponentMode.GridRowOnSelection
         cmdRCBdNew.Enabled = True
         cmdRCBdEdit.Enabled = True
        If flgChange = 1 Then
            cmdRCBdSave.Enabled = True
         Else
            cmdRCBdSave.Enabled = False
         End If
         cmdRCBdCancel.Enabled = True
         cmdRCBdDelete.Enabled = True
   End Select
End Sub

Private Sub Form_Load()
   Me.Width = 11325
   Me.Height = 7800
   Me.Top = 0
   Me.Left = 0
   Me.BackColor = MODULEBACKCOLOR
   Frame1(1).BackColor = Me.BackColor

   initialiseGrid
   RCSumTotal

   Call WheelHook(Me.hWnd)
End Sub

Private Sub initialiseGrid()
   Call ConfigureFlxBRMain
   flgEdit = 0
   flgChange = 0
   flgNew = 0
   Conn.Open getConnectionString
  
   LoadFlxRentBudgetDetails

   Call LoadFund
   
   Conn.Close
   Set Conn = Nothing
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Function TotalRCProperty(szPropertyID As String) As Double
   Dim Rst2 As New ADODB.Recordset

   If sModule = "RB" Then
      SQLStr = "SELECT GlobalRC.PropertyID, SUM(GlobalRC.TotalBudget) AS TOTALRENT " & _
               "From GlobalRC " & _
               "WHERE GlobalRC.PropertyID = '" & szPropertyID & "' " & _
               "GROUP BY GlobalRC.PropertyID;"
   Else
      SQLStr = "SELECT GlobalInsurance.PropertyID, SUM(GlobalInsurance.Amount) AS TOTALRENT " & _
               "From GlobalInsurance " & _
               "WHERE GlobalInsurance.PropertyID = '" & szPropertyID & "' " & _
               "GROUP BY GlobalInsurance.PropertyID;"
   End If
'Debug.Print SQLStr
   Rst2.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If Not Rst2.EOF Then
      TotalRCProperty = CDbl(Rst2!TOTALRENT)
   Else
      TotalRCProperty = 0
   End If

   Rst2.Close
   Set Rst2 = Nothing
End Function

Private Sub LoadFund()
   ' Error Handler
   On Error GoTo Error_Handler

   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim szSQL As String, adoRst As New ADODB.Recordset

   If sModule = "RB" Then
      szSQL = "SELECT FundID, FundName " & _
              "FROM Fund " & _
              "WHERE CategoryCode = 1;"
   Else
      szSQL = "SELECT FundID, FundName " & _
              "FROM Fund " & _
              "WHERE CategoryCode = 3;"
   End If

   adoRst.Open szSQL, Conn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "An insurance fund has not been setup for this property", , "N"
   Else
      ReDim Data(2, adoRst.RecordCount) As String

      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = Trim(adoRst.Fields.Item("FundID").Value)
         Data(1, rRow) = Trim(adoRst.Fields.Item("FundName").Value)
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
    
      cboRCFund.Clear
      cboRCFund.Column() = Data()
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading fund.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtBudget_GotFocus()
'   If flgEdit = 1 Or flgNew = 1 Then txtBudget.BackColor = &H80000013
End Sub

Private Sub txtBudget_LostFocus()
   computePpsf
'   If flgEdit = 1 Or flgNew = 1 Then txtBudget.BackColor = &H80000005
End Sub

Private Sub computePpsf()
 If Trim(txtBudget.text) <> "" And Trim(txtTotalArea.text) <> "" Then
      Dim ppsf As Double
      ppsf = CDbl(txtBudget.text) / CDbl(Trim(txtTotalArea.text))
      txtPpsf.text = FormatNumber(CDbl(ppsf), 2, , , vbDefault)
   End If
End Sub

Private Sub txtTotalArea_GotFocus()
'   If flgEdit = 1 Or flgNew = 1 Then txtTotalArea.BackColor = &H80000013
End Sub

Private Sub txtTotalArea_LostFocus()
   computePpsf
'   If flgEdit = 1 Or flgNew = 1 Then txtTotalArea.BackColor = &H80000005
End Sub

Private Sub updateGrid()
   If flgEdit = 0 Then
      flxRentBudgetDetails.AddItem ""
      If flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 0) <> "" Then flxRentBudgetDetails.AddItem ""
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 1) = frmGlobal1.cboProperty.BoundText
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 2) = CInt(cboRCFund.Value)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 3) = cboRCFund.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 5) = txtTotalArea.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 6) = txtPpsf.text
   Else
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 1) = frmGlobal1.cboProperty.BoundText
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 2) = CInt(cboRCFund.Value)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 3) = cboRCFund.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 5) = txtTotalArea.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 6) = txtPpsf.text
   End If

   flgEdit = 0
   flgNew = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   frmGlobal1.Enabled = True
'   frmGlobal1.SetFocus
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
