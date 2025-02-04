VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmServiceChargeRunFlag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Charge Year End"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7245
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4455
      TabIndex        =   2
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8730
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5775
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCBudgetDetails 
      Height          =   5025
      Left            =   90
      TabIndex        =   8
      Top             =   1035
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   8864
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
      Caption         =   "Year End Completed"
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   10
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client ID"
      Height          =   195
      Index           =   1
      Left            =   2745
      TabIndex        =   9
      Top             =   765
      Width           =   600
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Name"
      Height          =   195
      Index           =   0
      Left            =   6060
      TabIndex        =   7
      Top             =   795
      Width           =   825
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Area"
      Height          =   195
      Index           =   2
      Left            =   7530
      TabIndex        =   6
      Top             =   795
      Width           =   735
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price/SqFoot"
      Height          =   195
      Index           =   4
      Left            =   10485
      TabIndex        =   5
      Top             =   795
      Width           =   945
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Budget"
      Height          =   195
      Index           =   3
      Left            =   8895
      TabIndex        =   4
      Top             =   795
      Width           =   915
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Code"
      Height          =   195
      Index           =   5
      Left            =   4095
      TabIndex        =   3
      Top             =   765
      Width           =   780
   End
End
Attribute VB_Name = "frmServiceChargeRunFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadFlxSCBudgetDetails(adoconn As ADODB.Connection)
    '   On Error GoTo Err
    Dim i       As Integer
    Dim Rst     As New ADODB.Recordset
    Call ConfigFlxSCBudgetDetails
    '   szSQL = "SELECT g.*, f.FundName " & _
    '            "FROM GlobalSC g, Fund f " & _
    '            "WHERE CInt(g.Fund)=f.FundId;"
    'Modified by anol 20161013
    szSQL = "SELECT g.*, f.FundName,f.FundCode,P.clientID " & _
            "FROM GlobalSC g, Fund f,Property P " & _
            "WHERE CInt(g.Fund)=f.FundId AND P.propertyID=G.propertyID and F.CategoryCode=5 and P.clientID='" & frmServiceCharge.txtClientList.Tag & "';"
            
    Rst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

    With flxSCBudgetDetails
    i = 1
    If Not Rst.EOF Then
                ' txtSCBudgetTotal.text = "0"
                While Not Rst.EOF
                   .AddItem ""
                   .TextMatrix(i, 0) = Rst!budgetId
                   .TextMatrix(i, 1) = IIf(Rst!SCYearEndGenerated = True, "Yes", "No")
                   .TextMatrix(i, 2) = Rst!clientID
                   .TextMatrix(i, 3) = Rst!propertyID
                   .TextMatrix(i, 4) = Rst!Fund ' fund is fund Id in globalSC
                   .TextMatrix(i, 5) = Rst!FundCode
                   .TextMatrix(i, 6) = Rst!FundName
                   .TextMatrix(i, 7) = Rst!SCArea
                  ' txtSCBudgetTotal.text = Format(Rst!TotalBudget, "0.00") + Val(txtSCBudgetTotal.text)
                   .TextMatrix(i, 8) = Format(Rst!TotalBudget, "0.00") '
                   .TextMatrix(i, 9) = Format(Rst!ppsf, "0.00")
                   .TextMatrix(i, 10) = i - 1
                   .TextMatrix(i, 11) = IIf(IsNull(Rst!FinancialYear), "", Rst!FinancialYear)
                   '.RowHeight(i) = 0
                   i = i + 1
                   Rst.MoveNext
                Wend
      End If
       ' txtSCBudgetTotal.text = Format(txtSCBudgetTotal.text, "0.00")
      'initialiseMatrix
      .row = 0
      .col = 0
   End With

   Rst.Close
   Set Rst = Nothing
   Exit Sub
Err:
   MsgBox Err.description & "Error form LoadFlxSCBudgetDetails"
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

        Case TypeOf ctl Is PictureBox
          'PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
           bHandled = False

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
Private Sub ConfigFlxSCBudgetDetails()
   Dim szFlxHeader As String

   flxSCBudgetDetails.Rows = 1
   flxSCBudgetDetails.RowHeight(0) = 0
   flxSCBudgetDetails.Clear
   flxSCBudgetDetails.Cols = 11
   szFlxHeader$ = "BudgetID|PropertyID|<Fund|<FundCODE|>FundName|>TotalBudget|>SCArea|>PPSF|FY"
   flxSCBudgetDetails.FormatString = szFlxHeader$

   flxSCBudgetDetails.ColWidth(0) = 0
   flxSCBudgetDetails.ColWidth(1) = 1800
   flxSCBudgetDetails.ColWidth(2) = 1700 'Check box for run
   flxSCBudgetDetails.ColWidth(3) = 0 'THis can be used for clientID
   flxSCBudgetDetails.ColAlignment(3) = vbLeftJustify
   flxSCBudgetDetails.ColWidth(4) = 0
   flxSCBudgetDetails.ColWidth(5) = 2000 'lblRentCharges(2).Left - lblRentCharges(0).Left
    flxSCBudgetDetails.ColAlignment(5) = vbLeftJustify
   flxSCBudgetDetails.ColWidth(6) = lblRentCharges(2).Left - lblRentCharges(0).Left
    flxSCBudgetDetails.ColAlignment(6) = vbLeftJustify
   flxSCBudgetDetails.ColWidth(7) = lblRentCharges(3).Left - lblRentCharges(2).Left + 200
   flxSCBudgetDetails.ColWidth(8) = lblRentCharges(4).Left - lblRentCharges(3).Left - 200
   flxSCBudgetDetails.ColWidth(9) = lblRentCharges(4).Left - lblRentCharges(3).Left 'flxSCBudgetDetails.Width - lblRentCharges(4).Left - 300
   flxSCBudgetDetails.ColWidth(10) = 0
   flxSCBudgetDetails.ColWidth(11) = 0
   flxSCBudgetDetails.ColWidth(12) = 0

'   txtSCBudgetTotal.Width = flxSCBudgetDetails.ColWidth(7)
'   txtSCBudgetTotal.Left = lblRentCharges(2).Left
'
'   txtSCTotalArea.Width = flxSCBudgetDetails.ColWidth(6)
'   txtSCTotalArea.Left = lblRentCharges(3).Left
End Sub

Private Sub cmdCancel_Click()
    cmdSave.Enabled = False
    cmdEdit.Enabled = True
    flxSCBudgetDetails.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    flxSCBudgetDetails.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
End Sub

Private Sub cmdSCDBdClose_Click()
    
End Sub

Private Sub cmdSave_Click()
    cmdSave.Enabled = False
    flxSCBudgetDetails.Enabled = False
    Dim szSQL As String
    Dim iCounter As Integer
    If MsgBox("Do you want to save?", vbYesNo, "Please confirm") = vbNo Then Exit Sub
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    For iCounter = 1 To flxSCBudgetDetails.Rows - 1
        If flxSCBudgetDetails.TextMatrix(iCounter, 0) <> "" Then
            szSQL = "Update GlobalSC set SCYearEndGenerated=" & _
                IIf(flxSCBudgetDetails.TextMatrix(iCounter, 1) = "Yes", True, False) & "  where BudgetID='" & flxSCBudgetDetails.TextMatrix(iCounter, 0) & "'"
            adoconn.Execute szSQL
        End If
    Next
    Call Update_SC_Lease(adoconn)
    MsgBox "Service charge Year End has been updated."
    
    adoconn.Close
End Sub
Private Sub Update_SC_Lease(adoconn As ADODB.Connection)
   Dim Rst     As New ADODB.Recordset
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   
   
        On Error GoTo ErrHandler:
        
        'Resolved By BOSL.
        'Modified By Asif. Issue: 0000519. Date: 04-Jan-2015
        'Updating the service charge budgets through SQL rather than iteration which is time consuming.
        
        ' Charging Method: 2
        szSQL = "UPDATE LServiceCharges " & _
        "SET " & _
        "LServiceCharges.SCTotal = 0, " & _
        "LServiceCharges.SCAmount = 0 " & _
        "Where " & _
        "LServiceCharges.ChargingMethod = 2;"
        
        adoconn.Execute szSQL
        
'        szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
'        "Units AS U, Frequencies AS F,  Property AS P " & _
'        "SET " & _
'        "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
'        "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
'        "Where " & _
'        "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
'        "L.LeaseID = LServiceCharges.LeaseID AND " & _
'        "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
'        "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
'        "LServiceCharges.ChargingMethod = 2 and GSC.SCYearEndGenerated=false;"
'
'        adoconn.Execute szSQL
'
'        ' Charging Method: 4
'
'        szSQL = "UPDATE LServiceCharges " & _
'        "SET " & _
'        "LServiceCharges.SCTotal = 0, " & _
'        "LServiceCharges.SCAmount = 0, " & _
'        "LServiceCharges.CMFigure = 0 " & _
'        "Where " & _
'        "LServiceCharges.ChargingMethod = 4;"
'
'        adoconn.Execute szSQL
'
'        szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
'        "Units AS U, Frequencies AS F,  Property AS P " & _
'        "SET " & _
'        "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
'        "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
'        "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
'        "Where " & _
'        "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
'        "L.LeaseID = LServiceCharges.LeaseID AND " & _
'        "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
'        "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
'        "LServiceCharges.ChargingMethod = 4 and GSC.SCYearEndGenerated=false;"
'
'        adoconn.Execute szSQL

'******  fund  category not 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
    "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
    "Where FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 2 AND FN.CategoryCode<>5;"
    
    adoconn.Execute szSQL
    
    '******  fund  category = 5
     szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
    "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 2 and GSC.SCYearEndGenerated=false AND FN.CategoryCode=5;"
    
    adoconn.Execute szSQL
    
    
    
    ' Charging Method: 4
    
    szSQL = "UPDATE LServiceCharges " & _
    "SET " & _
    "LServiceCharges.SCTotal = 0, " & _
    "LServiceCharges.SCAmount = 0, " & _
    "LServiceCharges.CMFigure = 0 " & _
    "Where " & _
    "LServiceCharges.ChargingMethod = 4;"
    
    adoconn.Execute szSQL
    
    '******  fund  category not 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN  " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
    "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
    "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 4 AND FN.CategoryCode<>5;"
    
    adoconn.Execute szSQL
      '******  fund  category = 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P , Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
    "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
    "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 4 and GSC.SCYearEndGenerated=false AND FN.CategoryCode=5;"
    
    adoconn.Execute szSQL
    
        
        'MsgBox "The lease service charge budgets are updated successfully."

Exit Sub
ErrHandler:
   MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Could not update Service Charge Budget"
End Sub
Private Sub flxSCBudgetDetails_Click()
    If flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = "Yes" Then
        flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = "No"
        flxSCBudgetDetails.CellBackColor = vbWhite
        Exit Sub
    Else
        flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = "Yes"
    End If
End Sub

Private Sub Form_Load()
    Me.BackColor = MODULEBACKCOLOR
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Call LoadFlxSCBudgetDetails(adoconn)
    adoconn.Close
    flxSCBudgetDetails.Enabled = False
    Call WheelHook(Me.hWnd)
End Sub
