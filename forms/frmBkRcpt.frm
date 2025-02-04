VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBkRcpt 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4650
   ClientLeft      =   4050
   ClientTop       =   4230
   ClientWidth     =   5505
   Icon            =   "frmBkRcpt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5505
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   5295
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "Select"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   12648384
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   12632256
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "frmBkRcpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public szCaption As String

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSel_Click()
   frmPurchaseExpense.flxBankRcpt.TextMatrix(frmPurchaseExpense.flxBankRcpt.Row, frmPurchaseExpense.flxBankRcpt.Col) = flxSupplier.TextMatrix(flxSupplier.Row, flxSupplier.Col)
   If Me.Caption = "BANK" Then
      frmPurchaseExpense.flxBankRcpt.TextMatrix(frmPurchaseExpense.flxBankRcpt.Row, 11) = "T1"
      frmPurchaseExpense.nTaxCode = TaxRate(1)
   End If
   If Me.Caption = "TAX CODE" Then
      frmPurchaseExpense.nTaxCode = flxSupplier.TextMatrix(flxSupplier.Row, 1)
      frmPurchaseExpense.VatCalculation
   End If
   Unload Me
End Sub

Private Sub flxSupplier_DblClick()
   cmdSel_Click
End Sub

Private Sub Form_Load()
   Me.Top = frmPurchaseExpense.Top + frmPurchaseExpense.flxBankRcpt.Top + (285 * frmPurchaseExpense.flxBankRcpt.Rows) + 1000
   Me.Left = frmPurchaseExpense.iBRFlxLeft
   
   frmPurchaseExpense.tabPurExp.Enabled = False

   flxSupplier.ColWidth(0) = 1500
   flxSupplier.ColWidth(1) = 4500
   flxSupplier.ColWidth(2) = 10
   flxSupplier.ColWidth(3) = 10
   flxSupplier.TextMatrix(0, 0) = "A/C"
   flxSupplier.TextMatrix(0, 1) = "Name"
'*****************************************Delete following lines when no more needed*********
   flxSupplier.AddItem ""
   flxSupplier.TextMatrix(1, 0) = "TEST1"
   flxSupplier.TextMatrix(1, 1) = "SAGE DATA OBJECT ERROR"
   flxSupplier.AddItem ""
   flxSupplier.TextMatrix(2, 0) = "TEST2"
   flxSupplier.TextMatrix(2, 1) = "CAN NOT ACCESS SAGE DATAOBJ"
'*******************************************************************************************************

   Me.Caption = szCaption
   Select Case szCaption
   Case Is = "BANK"
      BankAccount
   Case Is = "NOMINAL LEDGER"
      NominalCode
   Case Is = "DEPARTMENTS"
      Deapartment
   Case Is = "COST CODE"
      CostCode
   Case Is = "UNIT"
      CustomerAccount
   Case Is = "TAX CODE"
      TaxCode
   Case Is = "PROJECT REF."
      ProjectRef
   End Select
End Sub

Private Sub ProjectRef()
   ' Error Handler
   On Error GoTo Error_Handler
   
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oProjects As SageDataObject120.Projects

   ' Declare Variables
   Dim szDataPath As String
   
   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
'   oSDO.Workspaces.Clear
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         ' Create Objects
         Set oProjects = oWS.CreateProjects
         flxSupplier.Clear
         
         If oProjects.Count = 0 Then
            MsgBox "No project code has been created", vbCritical, "Project Empty"
            GoTo Error_Handler
         End If
         
         Dim rRow As Integer
         For rRow = 1 To oProjects.Count
            flxSupplier.TextMatrix(rRow, 0) = CStr(oProjects.Item(rRow - 1).Reference)
            flxSupplier.TextMatrix(rRow, 1) = CStr(oProjects.Item(rRow - 1).Name)
         Next rRow

         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oProjects = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text
   
   Set oProjects = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmPurchaseExpense.Enabled = True
   frmPurchaseExpense.tabPurExp.Enabled = True
End Sub

Private Sub TaxCode()
   flxSupplier.ColWidth(0) = 1000
   flxSupplier.ColWidth(1) = 3000
   flxSupplier.TextMatrix(0, 0) = "CODE"
   flxSupplier.TextMatrix(0, 1) = "RATE"
   
   ' Error Handler
   On Error GoTo Error_Handler
   
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oTaxCode As SageDataObject120.ControlData

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
'   oSDO.Workspaces.Clear
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         Set oTaxCode = oWS.CreateObject("ControlData")
         flxSupplier.Clear
         Dim iTax As Integer
         For iTax = 0 To 99
            flxSupplier.TextMatrix(iTax + 1, 0) = "T" & iTax
            flxSupplier.TextMatrix(iTax + 1, 1) = CStr(oTaxCode.Fields.Item("T" & iTax & "_RATE").Value)
            flxSupplier.AddItem ""
         Next iTax
         
         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oTaxCode = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

'   MsgBox "The SDO generated the following error: " & oSDO.LastError.Text
   oWS.Disconnect
   Set oTaxCode = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
End Sub

Private Sub CustomerAccount()
   Dim rRow As Integer
   Dim Conn2 As New RDO.rdoConnection

   Dim szSql As String
   Dim rstRec As rdoResultset
   
   'Reset screen to show all the units in cboUnits.
   'Set the RDO Connections to the dataset
   Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn2.CursorDriver = rdUseIfNeeded
   Conn2.EstablishConnection rdDriverNoPrompt
   
   szSql = "SELECT UnitNumber, TenantCompanyName FROM Units ORDER BY UnitNumber"
   Set rstRec = Conn2.OpenResultset(szSql, rdOpenStatic, rdConcurReadOnly)

   If rstRec.EOF = False Then
      flxSupplier.Clear
      rRow = 1
      rstRec.MoveFirst
      flxSupplier.ColAlignment(0) = vbRightJustify
      While Not rstRec.EOF
         flxSupplier.TextMatrix(rRow, 0) = rstRec!UnitNumber
         If rstRec!TenantCompanyName <> "" Then
            flxSupplier.TextMatrix(rRow, 1) = rstRec!TenantCompanyName
         End If
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier.AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   Conn2.Close
End Sub

Private Sub CostCode()
   ' Error Handler
   On Error GoTo Error_Handler
   
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oProjCostCodes As SageDataObject120.ProjectCostCodes

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")
   
   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         ' Create Objects
         Set oProjCostCodes = oWS.CreateProjectCostCodes

         flxSupplier.Clear
         flxSupplier.TextMatrix(0, 0) = "CODE"
         flxSupplier.TextMatrix(0, 1) = "DESCRIPTION"

         Dim rRow As Integer
         For rRow = 1 To oProjCostCodes.Count
            flxSupplier.AddItem ""
            flxSupplier.TextMatrix(rRow, 0) = CStr(oProjCostCodes.Item(rRow - 1).Reference)
            flxSupplier.TextMatrix(rRow, 1) = CStr(oProjCostCodes.Item(rRow - 1).description)
         Next rRow

         flxSupplier.RemoveItem oProjCostCodes.Count + 1
         flxSupplier.RemoveItem oProjCostCodes.Count + 1

         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oProjCostCodes = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text
   
   Set oProjCostCodes = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
End Sub

Private Sub Deapartment()
   ' Error Handler
   On Error GoTo Error_Handler
   
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   '  Set oSDO = New SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   '  Set oWS = oSDO.Workspaces.Add("WkpsSupplier")
   Dim oDepartmentData As SageDataObject120.DepartmentData

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         Set oDepartmentData = oWS.CreateObject("DepartmentData")
         flxSupplier.Clear
         Dim rRow As Integer
         For rRow = 1 To oDepartmentData.Count
            oDepartmentData.Read (rRow)
            flxSupplier.TextMatrix(rRow, 0) = CStr(rRow)
            flxSupplier.TextMatrix(rRow, 1) = CStr(oDepartmentData.Fields.Item("NAME").Value)
            flxSupplier.AddItem ""
         Next rRow
         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

      MsgBox "The SDO generated the following error: " & oSDO.LastError.text
   Set oDepartmentData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub NominalCode()
' Error Handler
  On Error GoTo Error_Handler

  ' Declare Objects
  Dim oSDO As SageDataObject120.SDOEngine
  Dim oWS As SageDataObject120.Workspace
  Dim oNominalRecord As SageDataObject120.NominalRecord

  ' Declare Variables
  Dim szDataPath As String
  Dim bFlag As Boolean

  ' Create the SDO Engine object
  Set oSDO = New SageDataObject120.SDOEngine

  ' Create the workspace
  Set oWS = oSDO.Workspaces.Add("Example")

  ' Select Company.  The SelectCompany method takes the program install
  ' folder as a parameter
  szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
'  szDataPath = oSDO.SelectCompany("C:\Program Files\Sage\Accounts")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
   
      ' Try to Connect - Will throw an exception if it fails
      If oWS.Connect(szDataPath, "MANAGER", "", "Example") Then
   
        ' Create instance of NominalRecord object
        Set oNominalRecord = oWS.CreateObject("NominalRecord")
   
        ' Call the AddNew method
   '     oNominalRecord.AddNew
   
        oNominalRecord.MoveFirst
        Dim i As Integer
        i = oNominalRecord.Count
        For i = 1 To oNominalRecord.Count
           flxSupplier.TextMatrix(i, 0) = CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)
           flxSupplier.TextMatrix(i, 1) = CStr(oNominalRecord.Fields.Item("NAME").Value)
           flxSupplier.AddItem ""
           oNominalRecord.MoveNext
        Next i
        
        ' Disconnect
        oWS.Disconnect
      End If
    End If
   
    ' Destroy Objects
    Set oNominalRecord = Nothing
    Set oWS = Nothing
    Set oSDO = Nothing
   
    Exit Sub

' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text, vbOKOnly, "SDO Examples"
End Sub

Private Sub BankAccount()
   ' Error Handler
   On Error GoTo Error_Handler
   
   Dim clsBankAC As clsArray
   Dim iBankAc As Integer
   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oBankRecord As SageDataObject120.BankRecord
   Dim oNominalRecord As SageDataObject120.NominalRecord

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         Set oBankRecord = oWS.CreateObject("BankRecord")

         ' Move to the First Record
         oBankRecord.MoveFirst
         Set clsBankAC = New clsArray
         For iBankAc = 1 To oBankRecord.Count
            clsBankAC.AddItem oBankRecord.Fields.Item("ACCOUNT_REF").Value
            oBankRecord.MoveNext
         Next iBankAc

         Set oBankRecord = Nothing
         
         Set oNominalRecord = oWS.CreateObject("NominalRecord")

         oNominalRecord.MoveFirst
         
         Dim rRow As Integer
         Dim iRec As Integer
         rRow = 1
         For iRec = 1 To oNominalRecord.Count
            If clsBankAC.IsItem(CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)) Then
               flxSupplier.TextMatrix(rRow, 0) = CStr(oNominalRecord.Fields.Item("ACCOUNT_REF").Value)
               flxSupplier.TextMatrix(rRow, 1) = CStr(oNominalRecord.Fields.Item("NAME").Value)
               flxSupplier.AddItem ""
               rRow = rRow + 1
            End If
            oNominalRecord.MoveNext
         Next iRec
         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text
   
   Set oBankRecord = Nothing
   Set oNominalRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
End Sub
