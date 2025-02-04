VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{EC4A06C3-9499-4BA0-8D3C-6EDE133B1673}#1.1#0"; "HoverButtons.ocx"
Begin VB.Form frmExport2SAGE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posting to SAGE"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   Icon            =   "frmExort2SAGE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   5505
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00E9E8E4&
      Caption         =   "Select SAGE Transaction Date:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   15
      Top             =   760
      Width           =   3135
      Begin MSForms.OptionButton optDueDt 
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   540
         Width           =   1815
         VariousPropertyBits=   746588179
         BackColor       =   15329508
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "3201;423"
         Value           =   "0"
         Caption         =   "Due Date"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optIssueDt 
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1815
         VariousPropertyBits=   746588179
         BackColor       =   15329508
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "3201;423"
         Value           =   "0"
         Caption         =   "Issue Date"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optPostingDt 
         Height          =   280
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1335
         BackColor       =   15329508
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2355;494"
         Value           =   "0"
         Caption         =   "Posting Date"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CheckBox chkBPR 
      BackColor       =   &H00E9E8E4&
      Caption         =   "Bank Payments && Receipts"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   500
      TabIndex        =   3
      Top             =   3000
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CheckBox chkPI 
      BackColor       =   &H00E9E8E4&
      Caption         =   "Purchase Invoices && Credit Notes"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   500
      TabIndex        =   2
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Frame fraBackupQues 
      Caption         =   "Please confirm that - "
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CheckBox chkPrestigeUpdate 
         Caption         =   "I have taken a backup of Prestige."
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
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox chkSageUpdate 
         Caption         =   "I have taken a backup of SAGE."
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CheckBox chkReceipt 
      BackColor       =   &H00E9E8E4&
      Caption         =   "Receipts"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   500
      TabIndex        =   1
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin HoverButton.HoverControl cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   -2147483627
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancel"
   End
   Begin HoverButton.HoverControl cmdOK 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   -2147483627
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
   End
   Begin VB.CheckBox chkDemands 
      BackColor       =   &H00E9E8E4&
      Caption         =   "Demands"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   500
      TabIndex        =   0
      Top             =   480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFE0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   3615
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00E0FFE0&
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optExport 
         BackColor       =   &H00E0FFE0&
         Caption         =   "Post to SAGE"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
   End
   Begin MSForms.Frame Frame1 
      Height          =   615
      Left            =   120
      OleObjectBlob   =   "frmExort2SAGE.frx":08CA
      TabIndex        =   10
      Top             =   4800
      Width           =   3975
   End
   Begin MSForms.Frame Frame3 
      Height          =   615
      Left            =   120
      OleObjectBlob   =   "frmExort2SAGE.frx":12E2
      TabIndex        =   12
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Menu:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   1005
   End
   Begin MSForms.Image Image1 
      Height          =   3375
      Left            =   120
      Top             =   120
      Width           =   3975
      BackColor       =   15329508
      Size            =   "7011;5953"
   End
End
Attribute VB_Name = "frmExport2SAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iTotalTran As Integer

Private Sub chkBPR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Form_Unload 0
End Sub

Private Sub chkDemands_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Form_Unload 0
End Sub

Private Sub chkPI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Form_Unload 0
End Sub

Private Sub chkReceipt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Form_Unload 0
End Sub

Private Sub cmdOK_Click()
   MousePointer = vbHourglass

'PREVIEW BEFORE EXPROTED TO SAGE
   If optPreview.Value Then
      If chkDemands.Value Then                'DEMANDS
         ShowReport App.Path & szReportPath & "\PreViewExportedSAGE.rpt"
      End If
      If chkReceipt.Value Then                'RECEIPTS
         ShowReport App.Path & szReportPath & "\PreViewReceiptExportedSAGE.rpt"
      End If
      If chkPI.Value Then                     'PI & PC
         ShowReport App.Path & szReportPath & "\PreViewPIExportedSAGE.rpt"
      End If
      If chkBPR.Value Then                    'BP & BR
         ShowReport App.Path & szReportPath & "\PreViewBPnRExportedSAGE.rpt"
      End If
   End If
   If optExport.Value Then
'EXPORTED TO SAGE
      iTotalTran = 0
      If chkSageUpdate.Value And chkPrestigeUpdate.Value Then
         If chkDemands.Value Then            'DEMANDS
            Call ExportSI2Sage
         End If
         If chkReceipt.Value Then            'RECEIPT
            Call ExportSR2Sage
         End If
         If chkPI.Value Then                 'PI & PC
            Call ExportPI2Sage
         End If
         If chkBPR.Value Then                'BP & BR
            Call ExportBPnR2Sage
         End If
         fraBackupQues.Visible = False
         chkSageUpdate.Value = False
         chkPrestigeUpdate.Value = False
         optPreview.Value = True
      Else
         If chkSageUpdate.Value = 0 Then
            MsgBox "Please take a backup of your SAGE data.", vbInformation, "Data backup"
         End If
         If chkPrestigeUpdate.Value = 0 Then
            MsgBox "Please take a backup of your Prestige data.", vbInformation, "Data backup"
         End If
      End If
   End If
   MousePointer = vbDefault
End Sub

Private Sub ExportBPnR2Sage()
   Dim iCount As Integer
   Dim szaBPnR_ID() As String
'
'  Export data to sage from demand and demandsplit table
'  Error Handler
   On Error GoTo Error_Handler
'
'    Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oTransactionPost As SageDataObject120.TransactionPost
   Dim oSplitData As SageDataObject120.SplitData

'   Declare Variables for Database connectivity
   Dim adoConn As ADODB.Connection
   Dim adoRstBPnR As ADODB.Recordset

'   Declare Variables
   Dim szDataPath As String, szSQLStr As String

'   Set the connection to the databases
   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

'   Connect to Demands table to add new demands.
   Set adoRstBPnR = New ADODB.Recordset
   szSQLStr = "SELECT MY_ID, TRAN_ID, BANK_AC, TRAN_DATE, TRANS, TRAN_TYPE, " & _
                  "UNIT_ID, NOMINAL_CODE, DEPT_ID, PROJ_REF, COST_CODE, DESCRIPTION, " & _
                  "NET_AMOUNT, TAX_CODE, VAT, UPDATE_SAGE, RECHARGED, RECHARABLE, " & _
                  "TransactionType, SAGEPostingNumber " & _
              "FROM tlbBankPayment " & _
              "WHERE UPDATE_SAGE = FALSE;"

   adoRstBPnR.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   if there no data to be exported then exit from this method
   If adoRstBPnR.EOF Then
      MsgBox "There no Purchase Invoice to be exported to SAGE.", vbOKOnly, "SAGE"
      chkBPR.Value = False
      Set adoConn = Nothing
      Set adoRstBPnR = Nothing
      Exit Sub
   End If

'   Create the SDO Engine Object
   Set oSDO = New SageDataObject120.SDOEngine
'
'    Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Prestige")

'    Select Company.  The SelectCompany method takes the program install folder as a parameter
'   read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
'       Select Company. The SelectCompany method takes the program install folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If

'      Try to Connect - Will throw an exception if it fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then

'        Populate Header fields
       iCount = 0
'Note:
'        The ACCOUNT_REF field must be populated with a valid
'        Customer Account Reference
    While Not adoRstBPnR.EOF
'        Create Instances of Objects
       Set oTransactionPost = oWS.CreateObject("TransactionPost")

'        Populate Header Fields
       oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstBPnR.Fields("BANK_AC").Value)
       oTransactionPost.Header("Date").Value = CDate(adoRstBPnR.Fields("TRAN_DATE").Value)
       oTransactionPost.Header("Posted_Date").Value = CDate(Date)
       oTransactionPost.Header("TYPE").Value = CByte(adoRstBPnR.Fields("TransactionType").Value)
       oTransactionPost.Header("DETAILS").Value = CStr(adoRstBPnR.Fields("DESCRIPTION").Value)
       oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstBPnR.Fields("TRAN_ID").Value)
       oTransactionPost.Header("BANK_CODE").Value = CStr(adoRstBPnR.Fields("BANK_AC").Value)
       oTransactionPost.Header("EURO_GROSS").Value = CDbl(adoRstBPnR.Fields("NET_AMOUNT").Value + adoRstBPnR.Fields("VAT").Value)
       oTransactionPost.Header("EURO_RATE").Value = CDbl(1)
       oTransactionPost.Header("FOREIGN_GROSS").Value = CDbl(adoRstBPnR.Fields("NET_AMOUNT").Value + adoRstBPnR.Fields("VAT").Value)
       oTransactionPost.Header("FOREIGN_RATE").Value = CDbl(1)
       oTransactionPost.Header("INTEREST_RATE").Value = CDbl(0)
       oTransactionPost.Header("NET_AMOUNT").Value = CDbl(adoRstBPnR.Fields("NET_AMOUNT").Value)

'        Note:
'        The PI Transaction have 1 split
'        Add a split to the Header's Item collection
       Set oSplitData = oTransactionPost.Items.Add()

'        Fill in the Split Fields
       oSplitData.Fields.Item("TYPE").Value = CByte(adoRstBPnR.Fields("TransactionType").Value)
'oSplitData.Fields.Item("DEPT_NUMBER").Value = CByte(????)
       oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstBPnR.Fields("NOMINAL_CODE").Value)
       oSplitData.Fields.Item("TAX_CODE").Value = CInt(Mid(adoRstBPnR.Fields("TAX_CODE").Value, 2))
       oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstBPnR.Fields("NET_AMOUNT").Value)
       oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(adoRstBPnR.Fields("VAT").Value)
       oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
       oSplitData.Fields.Item("DETAILS").Value = CStr(oTransactionPost.Header("DETAILS").Value)

       If oTransactionPost.Update Then
          SavePostingNumber "tlbBankPayment", "MY_ID", adoRstBPnR.Fields("MY_ID").Value, oTransactionPost.PostingNumber, adoConn
          iCount = iCount + 1
          ReDim Preserve szaBPnR_ID(iCount) As String
          szaBPnR_ID(iCount - 1) = CStr(adoRstBPnR.Fields("MY_ID").Value)
       End If
       adoRstBPnR.MoveNext
       Set oSplitData = Nothing
       Set oTransactionPost = Nothing
    Wend
    adoRstBPnR.Close

    Dim i As Integer
    For i = 0 To iCount - 1
       ' Update the record in header table, UPDATE_SAGE = True
       adoRstBPnR.Open "UPDATE tlbBankPayment " & _
                        "SET UPDATE_SAGE = TRUE " & _
                        "WHERE UPDATE_SAGE = FALSE AND " & _
                              "MY_ID = '" & szaBPnR_ID(i) & "'", adoConn, adOpenStatic, adLockReadOnly
    Next i
    MsgBox iCount & " BP & BR Transactions Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
'     Disconnect SAGE
    oWS.Disconnect
   End If

   Set oWS = Nothing
   Set oSDO = Nothing

   Set adoRstBPnR = Nothing
   Set adoConn = Nothing

   Exit Sub

'    Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text, vbOKOnly, "Posted to SAGE"
   ' Destroy Objects
   Set oTransactionPost = Nothing
   Set oSplitData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
   Set adoConn = Nothing
   Set adoRstBPnR = Nothing
End Sub

Private Sub ExportPI2Sage()
   Dim iCount As Integer
   Dim szaPI_ID() As String
'
'  Export data to sage from demand and demandsplit table
'  Error Handler
'   On Error GoTo Error_Handler
'
'    Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
'   Dim oSalesRecord As SageDataObject120.SalesRecord
   Dim oTransactionPost As SageDataObject120.TransactionPost
   Dim oSplitData As SageDataObject120.SplitData

'   Declare Variables for Database connectivity
   Dim adoConn As ADODB.Connection
   Dim adoRstPI As ADODB.Recordset

'   Declare Variables
   Dim szDataPath As String, szSQLStr As String

'   Set the connection to the databases
   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

'   Connect to Demands table to add new demands.
   Set adoRstPI = New ADODB.Recordset
   szSQLStr = "SELECT * " & _
              "FROM tlbPurchaseInvoice " & _
              "WHERE UPDATE_SAGE = FALSE;"

   adoRstPI.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   if there no data to be exported then exit from this method
   If adoRstPI.EOF Then
      MsgBox "There no Purchase Invoice to be exported to SAGE.", vbOKOnly, "SAGE"
      chkPI.Value = False
      Set adoConn = Nothing
      Set adoRstPI = Nothing
      Exit Sub
   End If

'   Set adoRstPI = New ADODB.Recordset

'   Create the SDO Engine Object
   Set oSDO = New SageDataObject120.SDOEngine

'    Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Prestige")

'    Select Company.  The SelectCompany method takes the program install
'    folder as a parameter
   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
'       Select Company. The SelectCompany method takes the program install
'       folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If
'      Try to Connect - Will throw an exception if it fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
'
'        Populate Header fields
'
       iCount = 0
'Note:
'        The ACCOUNT_REF field must be populated with a valid
'        Customer Account Reference
    While Not adoRstPI.EOF
'        Create Instances of Objects
       Set oTransactionPost = oWS.CreateObject("TransactionPost")

       oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstPI.Fields("SUPP_AC").Value)
       oTransactionPost.Header("Date").Value = CDate(adoRstPI.Fields("TRAN_DATE").Value)
       oTransactionPost.Header("Posted_Date").Value = CDate(Date)
       oTransactionPost.Header("TYPE").Value = CByte(adoRstPI.Fields("TransactionType").Value)
       oTransactionPost.Header("DETAILS").Value = CStr(adoRstPI.Fields("DESCRIPTION").Value)
       oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstPI.Fields("INV_NO").Value)
       oTransactionPost.Header("EURO_GROSS").Value = CDbl(adoRstPI.Fields("NET_AMOUNT").Value + adoRstPI.Fields("VAT").Value)
       oTransactionPost.Header("EURO_RATE").Value = CDbl(1)
       oTransactionPost.Header("FOREIGN_GROSS").Value = CDbl(adoRstPI.Fields("NET_AMOUNT").Value + adoRstPI.Fields("VAT").Value)
       oTransactionPost.Header("FOREIGN_RATE").Value = CDbl(1)
       oTransactionPost.Header("INTEREST_RATE").Value = CDbl(0)
       oTransactionPost.Header("NET_AMOUNT").Value = CDbl(adoRstPI.Fields("NET_AMOUNT").Value)

'        Note:
'        The PI Transaction have 1 split
'        Add a split to the Header's Item collection
       Set oSplitData = oTransactionPost.Items.Add

'        Populate Split Fields
       oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)

       If adoRstPI.Fields("DEPT_ID").Value <> "" Then _
          oSplitData.Fields.Item("DEPT_NUMBER").Value = CInt(adoRstPI.Fields("DEPT_ID").Value)

       oSplitData.Fields.Item("DETAILS").Value = CStr(adoRstPI.Fields("DESCRIPTION").Value)  'data inserting: 'Service charge description'
       oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstPI.Fields("NET_AMOUNT").Value) 'data inserting: 1000
       oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstPI.Fields("NOMINAL_CODE").Value)  'data inserting: 800002
       oSplitData.Fields.Item("POSTED_DATE").Value = CDate(oTransactionPost.Header("POSTED_DATE").Value)
       oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(adoRstPI.Fields("VAT").Value)   'data inserting: 175
       oSplitData.Fields.Item("TAX_CODE").Value = CInt(Mid(adoRstPI.Fields("TAX_CODE").Value, 2))  'data inserting: 1
       oSplitData.Fields.Item("TYPE").Value = CByte(oTransactionPost.Header("TYPE").Value)

       If oTransactionPost.Update Then
          SavePostingNumber "tlbPurchaseInvoice", "MY_ID", adoRstPI.Fields("MY_ID").Value, oTransactionPost.PostingNumber, adoConn
          iCount = iCount + 1
          ReDim Preserve szaPI_ID(iCount) As String
          szaPI_ID(iCount - 1) = CStr(adoRstPI.Fields("MY_ID").Value)
       End If
       adoRstPI.MoveNext
       Set oSplitData = Nothing
       Set oTransactionPost = Nothing
    Wend
    adoRstPI.Close

    Dim i As Integer
    For i = 0 To iCount - 1
       ' Update the record in header table, UPDATE_SAGE = True
       adoRstPI.Open "UPDATE tlbPurchaseInvoice " & _
                        "SET UPDATE_SAGE = TRUE " & _
                        "WHERE UPDATE_SAGE = FALSE AND " & _
                              "MY_ID = '" & szaPI_ID(i) & "'", adoConn, adOpenStatic, adLockReadOnly
    Next i
    MsgBox iCount & " PI Transactions Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
'     Disconnect SAGE
    oWS.Disconnect
   End If
'
   Set oWS = Nothing
   Set oSDO = Nothing
'
   Set adoRstPI = Nothing
'   Set adoRstPI = Nothing
   Set adoConn = Nothing
'
   Exit Sub
'
'    Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text, vbOKOnly, "Posted to SAGE"
   ' Destroy Objects
   Set oTransactionPost = Nothing
   Set oSplitData = Nothing
'   Set oSalesRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
   
   Set adoConn = Nothing
   Set adoRstPI = Nothing
End Sub

Private Sub SetMark(ByVal bRef As Boolean, ByVal bSucc As Boolean, ByVal cOS_Balance As Currency, ByVal tranID As Long, adoConn As ADODB.Connection)
   Dim szSQLStr As String, szMark As String
   Dim adoRptCh As New ADODB.Recordset

   If bSucc Then
      szMark = "S"
   Else
      If bRef Then
         szMark = "F" & CStr(cOS_Balance)
      Else
         szMark = "F-1"
      End If
   End If

   szSQLStr = "SELECT UpDateSage, spare2  " & _
              "FROM tlbReceipt AS RPT " & _
              "WHERE RPT.TransactionID = " & tranID & ";"

   adoRptCh.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   adoRptCh.Fields.Item("UpDateSage").Value = IIf(szMark = "S", False, True)
   adoRptCh.Fields.Item("spare2").Value = szMark
   adoRptCh.Update

   adoRptCh.Close
   Set adoRptCh = Nothing

   iTotalTran = iTotalTran + 1
End Sub

' This method will check all invoices' outstanding amount with the receipt amount.
' If receipt amount > outstanding amount then this method will mark the receipt transaction
' not to exproted to sage. because sage does not accept it.
'
' If the receipt transaction not prossible to update then it will be marked at "spare2" field,
' in tlbReceipt table. If its possible to update then it will mark "S" into "spare2" field. Otherwise
' it will mark by "F" + outstanding amount in SAGE of the invoice (ex. F100).
Private Sub JustifyOSAmtOfRptToExport(adoRpt As ADODB.Recordset, adoConn As ADODB.Connection, ByVal oWS As SageDataObject120.Workspace)
   Dim szSQLStr As String, lInvoiceHeader As Long, bSucc As Boolean, j As Integer
   Dim cOS_Balance  As Currency, bRef As Boolean
   Dim oHeaderData As SageDataObject120.HeaderData
   Dim oSplitData As SageDataObject120.SplitData
   Dim adoRptCh As New ADODB.Recordset

   While Not adoRpt.EOF
      szSQLStr = "SELECT DR.SageAccountNumber AS SAN, DR.SAGEPostingNumber AS SPN, " & _
                     "RPT.RDate, RPT.ReceiptAmount, RPT.TransactionID, " & _
                     "DSR.TotalAmount, DSR.Description, RPT.Ref, RPT.Type  "
      szSQLStr = szSQLStr + _
                 "FROM tlbReceipt AS RPT, DemandRecords AS DR, DemandSplitRecords AS DSR "
      szSQLStr = szSQLStr + _
                 "WHERE RPT.IsSageUpdate = True And " & _
                     "RPT.UpDateSage = False And " & _
                     "RPT.DemandRef = DSR.DSR And " & _
                     "DSR.DemandID = DR.DemandID AND DR.SAGEPostingNumber <> 0 AND " & _
                     "DR.SageAccountNumber = '" & adoRpt!SageAccountNumber & "' AND " & _
                     "RPT.RDate = #" & Format(adoRpt!RDate, "dd mmmm yyyy") & "#;"
'Debug.Print szSQLStr
      adoRptCh.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRptCh.EOF
         If adoRptCh!Type <> "4" Then
            Set oHeaderData = oWS.CreateObject("HeaderData")
            lInvoiceHeader = CLng(adoRptCh!SPN)
            oHeaderData.Read (lInvoiceHeader)

            Set oSplitData = oHeaderData.Link

            oSplitData.MoveFirst

            bSucc = False
            cOS_Balance = -1
            For j = 1 To oSplitData.Count
               If oSplitData.Fields.Item("INTERNAL_REF").Value = "" Or IsNull(oSplitData.Fields.Item("INTERNAL_REF").Value) Then
                  bRef = False
                  If Val(oSplitData.Fields.Item("Net_Amount").Value) + Val(oSplitData.Fields.Item("Tax_Amount").Value) - Val(oSplitData.Fields.Item("AMOUNT_PAID").Value) >= Val(adoRptCh!ReceiptAmount) And _
                        InStr(oSplitData.Fields.Item("Details").Value, adoRptCh!description) > 0 Then
                     bSucc = True
                     Exit For
                  End If
               Else
                  If oSplitData.Fields.Item("INTERNAL_REF").Value = adoRptCh!Ref Then
                     bRef = True
                     If Val(oSplitData.Fields.Item("Net_Amount").Value) + Val(oSplitData.Fields.Item("Tax_Amount").Value) - Val(oSplitData.Fields.Item("AMOUNT_PAID").Value) >= Val(adoRptCh!ReceiptAmount) Then
                        bSucc = True
                     Else
                        bSucc = False
                        cOS_Balance = Val(oSplitData.Fields.Item("Net_Amount").Value) + Val(oSplitData.Fields.Item("Tax_Amount").Value) - Val(oSplitData.Fields.Item("AMOUNT_PAID").Value)
                     End If
                     Exit For
                  End If
               End If
               oSplitData.MoveNext
            Next j

            Call SetMark(bRef, bSucc, cOS_Balance, adoRptCh!TransactionID, adoConn)
         Else
            Call SetMark(True, True, 0, adoRptCh!TransactionID, adoConn)
         End If

         adoRptCh.MoveNext
      Wend
      adoRptCh.Close
      adoRpt.MoveNext
   Wend

   Set adoRptCh = Nothing

  ' Destroy the Objects
   Set oSplitData = Nothing
   Set oHeaderData = Nothing
End Sub

Private Sub ExportSR2Sage()
   Dim iCount As Integer, szSQLStr As String
   Dim i As Integer, iSplit As Integer, j As Integer
   Dim laTransID() As Long

   On Error GoTo Error_Handler

'   Declare Variables for Database connectivity
   Dim adoConn       As New ADODB.Connection
   Dim adoRstRpt     As New ADODB.Recordset
   Dim adoRstRptCh   As New ADODB.Recordset
   Dim adoPoA        As New ADODB.Recordset

'   Set the connection to the databases
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

   szSQLStr = "SELECT RPT.SageAccountNumber, RPT.RDate, RPT.TYPE, " & _
                  "RPT.BankCode, SUM(RPT.ReceiptAmount) AS RAMT, RPT.Spare1 "
   szSQLStr = szSQLStr + _
              "FROM tlbReceipt AS RPT "
   szSQLStr = szSQLStr + _
              "WHERE RPT.IsSageUpdate=True And " & _
                  "RPT.UpDateSage=False "
   szSQLStr = szSQLStr + _
              "GROUP BY  RPT.SageAccountNumber, RPT.RDate, RPT.TYPE, RPT.BankCode, RPT.Spare1;"
'Debug.Print szSQLStr
   adoRstRpt.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly

   szSQLStr = "SELECT tblPoA.* " & _
              "FROM tblPoA " & _
              "WHERE IsSageUpdate = True And " & _
                  "UpDateSage = False;"
'Debug.Print szSQLStr
   adoPoA.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly

   If adoRstRpt.EOF And adoPoA.EOF Then
      MsgBox "There are no transaction to update to SAGE.", vbInformation + vbOKOnly, "Receipt & Payment on Account"
      chkReceipt.Value = False
      adoRstRpt.Close
      adoPoA.Close
      adoConn.Close
      Set adoRstRpt = Nothing
      Set adoPoA = Nothing
      Set adoConn = Nothing
      Exit Sub
   End If

   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oTransactionPost As SageDataObject120.TransactionPost
   Dim oHeaderData As SageDataObject120.HeaderData
   Dim oSplitData As SageDataObject120.SplitData

   ' Declare Variables
   Dim szDataPath As String, iCtr As Integer
   Dim bFlag As Boolean, bBreak As Boolean
   Dim lInvoiceHeader As Long, lReceiptHeader As Long, lReceiptSplit As Long
   Dim lTotalAmount As Double, lAmountLeft As Double

  ' Create the SDO Engine Object
   Set oSDO = New SageDataObject120.SDOEngine

'   Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Prestige")

'   Select Company.  The SelectCompany method takes the program install
'   folder as a parameter
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)

'   A U.I. for company selection is presented to the user. If a company is selected,
'   the path will be passed to the szDataPath variable.
'   If not, or the Cancel button is selected, the variable will be left empty.
   If szDataPath <> "" Then
'     Try to Connect - Will throw an exception if it fails
      If Not oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
         MsgBox "The Prestige has failed to create a connection with SAGE. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Connection with SAGE Failed"

         adoRstRpt.Close
         adoPoA.Close
         adoConn.Close

'        Destroy the Objects
         Set oWS = Nothing
         Set oSDO = Nothing

         Set adoRstRpt = Nothing
         Set adoPoA = Nothing
         Set adoConn = Nothing
         Exit Sub
      End If
   Else
      MsgBox "There are some problems with the SAGE configuration in the system registry. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Registry Error"

      adoRstRpt.Close
      adoConn.Close

'     Destroy the Objects
      Set oWS = Nothing
      Set oSDO = Nothing

      Set adoRstRpt = Nothing
      Set adoConn = Nothing
      Exit Sub
   End If

   If adoRstRpt.EOF Then
      MsgBox "There are no Receipt to update to SAGE.", vbInformation + vbOKOnly, "Receipt"
      adoRstRpt.Close
      Set adoRstRpt = Nothing
      GoTo PoA
   End If

   JustifyOSAmtOfRptToExport adoRstRpt, adoConn, oWS
   adoRstRpt.Close
   ReDim Preserve laTransID(iTotalTran - 1) As Long

'   Main code segment to export to SAGE

   szSQLStr = "SELECT RPT.SageAccountNumber, RPT.RDate, RPT.TYPE, " & _
                  "RPT.BankCode, SUM(RPT.ReceiptAmount) AS RAMT, " & _
                  "RPT.Spare1, RPT.Spare3, RPT.Spare4 "
   szSQLStr = szSQLStr + _
              "FROM tlbReceipt AS RPT "
   szSQLStr = szSQLStr + _
              "WHERE RPT.IsSageUpdate = True And " & _
                  "RPT.UpDateSage = False And " & _
                  "(isnull(RPT.spare2) OR LEFT(RPT.spare2,1) = 'S') "
   szSQLStr = szSQLStr + _
              "GROUP BY  RPT.SageAccountNumber, RPT.RDate, RPT.TYPE, " & _
                  "RPT.BankCode, RPT.Spare1, RPT.Spare3, RPT.Spare4;"
'Debug.Print szSQLStr
   adoRstRpt.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly

   iCount = 0
   While Not adoRstRpt.EOF
      If adoRstRpt.Fields.Item("spare3").Value = "" Or IsNull(adoRstRpt.Fields.Item("spare3").Value) Then
'         Create an instance of TransactionPost for the Sales Receipt                  #SR#
         Set oTransactionPost = oWS.CreateObject("TransactionPost")

   '      Fill in the Header fields
         oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstRpt!SageAccountNumber)
         oTransactionPost.Header("DATE").Value = CDate(adoRstRpt!RDate)
         oTransactionPost.Header("POSTED_DATE").Value = CDate(Date)
         oTransactionPost.Header("TYPE").Value = CByte(sdoSR)                    'Sales Receipt -> 3
         oTransactionPost.Header("DETAILS").Value = "Sales Receipt"
         oTransactionPost.Header("BANK_CODE").Value = CStr(adoRstRpt!BankCode)
         oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstRpt!Spare1)

   '      Create a split item by adding an empty split to the Items
   '      collection of the TransactionPost Object.
         Set oSplitData = oTransactionPost.Items.Add()

   '      Fill in the Split fields - note a Sales Receipt only has one split
         oSplitData.Fields.Item("TYPE").Value = CByte(sdoSR)                   'Sales Receipt -> 3
         oSplitData.Fields.Item("DEPT_NUMBER").Value = CStr(adoRstRpt!spare4)
         oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstRpt!BankCode)
         oSplitData.Fields.Item("TAX_CODE").Value = CInt(9)
         oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstRpt!RAMT)
         oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(0)
         oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
         oSplitData.Fields.Item("DETAILS").Value = CStr(oTransactionPost.Header("DETAILS").Value)

   '      Update the TransactionPost Object
         If oTransactionPost.Update Then
            lReceiptHeader = oTransactionPost.PostingNumber

            Set oHeaderData = oWS.CreateObject("HeaderData")
            oHeaderData.Read (lReceiptHeader)
            lReceiptSplit = oHeaderData.Fields.Item("FIRST_SPLIT").Value

            Set oTransactionPost = Nothing
            Set oSplitData = Nothing
   
   'Finding the SAGE Posting Number of the Invoice from Demand Record to allocate the receipt against that Invoice
            szSQLStr = "SELECT DR.SageAccountNumber AS SAN, DR.SAGEPostingNumber AS SPN, " & _
                           "RPT.RDate, RPT.ReceiptAmount, RPT.TransactionID, " & _
                           "DSR.TotalAmount, DSR.Description, RPT.Type "
            szSQLStr = szSQLStr + _
                       "FROM tlbReceipt AS RPT, DemandRecords AS DR, DemandSplitRecords AS DSR "
            szSQLStr = szSQLStr + _
                       "WHERE RPT.Type = " & adoRstRpt!Type & " And RPT.IsSageUpdate = True And " & _
                           "RPT.UpDateSage = False And " & _
                           "RPT.DemandRef = DSR.DSR And " & _
                           "DSR.DemandID = DR.DemandID AND DR.SAGEPostingNumber <> 0 AND " & _
                           "DR.SageAccountNumber = '" & adoRstRpt!SageAccountNumber & "';"
   'Debug.Print szSQLStr
            adoRstRptCh.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
   
            While Not adoRstRptCh.EOF
               If adoRstRptCh.Fields.Item("Type").Value <> CByte(sdoSA) Then
                  lInvoiceHeader = CLng(adoRstRptCh!SPN)             'SPN -> SAGE Posting Number
                  oHeaderData.Read (lInvoiceHeader)
   
                  Set oSplitData = oHeaderData.Link
   
                  oSplitData.MoveFirst
   
                  bBreak = False
                  For j = 1 To oSplitData.Count
                     If Val(oSplitData.Fields.Item("Net_Amount").Value) + _
                           Val(oSplitData.Fields.Item("Tax_Amount").Value) = Val(adoRstRptCh!TotalAmount) And _
                           InStr(oSplitData.Fields.Item("Details").Value, adoRstRptCh!description) > 0 Then
                        bBreak = True
                     End If
                     If Not bBreak Then
                        oSplitData.MoveNext
                     Else
                        Exit For
                     End If
                  Next j
   
                  Set oTransactionPost = oWS.CreateObject("TransactionPost")
   
                  If oTransactionPost.AllocatePayment(CLng(oSplitData.RecordNumber), CLng(lReceiptSplit), _
                        CDbl(adoRstRptCh!ReceiptAmount), CDate(adoRstRptCh!RDate)) Then
                     oSplitData.MoveNext
                     laTransID(iCount) = adoRstRptCh!TransactionID
                     iCount = iCount + 1
                  End If
               Else                       'PAYMENT ON ACCOUNT - DOES NOT ALLOCATE AGAINST ANY TRANSACTION
                  laTransID(iCount) = adoRstRptCh!TransactionID
                  iCount = iCount + 1
               End If
               adoRstRptCh.MoveNext
            Wend
            adoRstRptCh.Close
         End If
      End If
      adoRstRpt.MoveNext
   Wend
   adoRstRpt.Close

' ***********************************************************************************************************
' *********************** ALLOCATION OF CREDIT AGAINST INVOICE **********************************************
' ***********************************************************************************************************
'  Exporting Allocations of 'Credit Invoice' against Invoice
   Dim lSPN_PoA As Long, lSPN_Inv As Long, bAllocated As Boolean, lSPN_CN As Long
   Dim oCreditHeaderData As SageDataObject120.HeaderData, oCreditSplitData As SageDataObject120.SplitData
   Dim oInvHeaderData As SageDataObject120.HeaderData, oInvSplitData As SageDataObject120.SplitData

   szSQLStr = "SELECT RPT.TransactionID, RPT.SageAccountNumber, RPT.RDate, RPT.TYPE, " & _
                  "RPT.BankCode, RPT.Amount as Amt, RPT.ReceiptAmount AS RAmt, RPT.Details, " & _
                  "RPT.Spare1, RPT.Ref, RPT.Allocation, RPT.Spare3 "
   szSQLStr = szSQLStr + _
              "FROM tlbReceipt AS RPT "
   szSQLStr = szSQLStr + _
              "WHERE RPT.IsSageUpdate = True And " & _
                  "RPT.UpDateSage = False And " & _
                  "(isnull(RPT.spare2) OR LEFT(RPT.spare2,1) = 'S');"
'Debug.Print szSQLStr
   adoRstRpt.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly

   Set oCreditHeaderData = oWS.CreateObject("HeaderData")
   Set oInvHeaderData = oWS.CreateObject("HeaderData")

'   iCount = 0
   While Not adoRstRpt.EOF
      If adoRstRpt.Fields.Item("spare3").Value <> "" Then
         If Val(adoRstRpt.Fields.Item("spare3").Value) = 4 Then                  'Allocating PoA
            lSPN_PoA = SagePostingNumberPoA(adoRstRpt.Fields.Item("Allocation").Value, adoConn)
            oCreditHeaderData.Read (lSPN_PoA)
   
            lSPN_Inv = SagePostingNumberInv(adoRstRpt.Fields.Item("TransactionID").Value, adoConn)
            oInvHeaderData.Read (lSPN_Inv)
   
            'Link to the Credit Splits
            Set oCreditSplitData = oCreditHeaderData.Link
            oCreditSplitData.MoveFirst
   
            Set oInvSplitData = oInvHeaderData.Link
            oInvSplitData.MoveFirst
   
            For i = 1 To oInvSplitData.Count
               If Val(oInvSplitData.Fields.Item("NET_AMOUNT").Value) - _
                     Val(oInvSplitData.Fields.Item("AMOUNT_PAID").Value) >= _
                     Val(adoRstRpt.Fields.Item("RAmt").Value) Then
                  Exit For
               End If
               oInvSplitData.MoveNext
            Next i
            If i <= oInvSplitData.Count Then
               Set oTransactionPost = oWS.CreateObject("TransactionPost")
               bAllocated = oTransactionPost.AllocatePayment(CInt(oInvSplitData.RecordNumber), _
                              CInt(oCreditSplitData.RecordNumber), CDbl(adoRstRpt.Fields.Item("RAmt").Value), _
                              CDate(adoRstRpt.Fields.Item("RDate").Value))
               If bAllocated Then
                  laTransID(iCount) = adoRstRpt.Fields.Item("TransactionID").Value
                  iCount = iCount + 1
               End If
            End If
            Set oCreditSplitData = Nothing
            Set oInvSplitData = Nothing
         End If
   
         If Val(adoRstRpt.Fields.Item("spare3").Value) = 2 Then                  'Allocating Credit Note
            lSPN_CN = SagePostingNumberInv(adoRstRpt.Fields.Item("Allocation").Value, adoConn)
            oCreditHeaderData.Read (lSPN_CN)
   
            lSPN_Inv = SagePostingNumberInv(adoRstRpt.Fields.Item("TransactionID").Value, adoConn)
            oInvHeaderData.Read (lSPN_Inv)
   
            'Link to the Credit Splits
            Set oCreditSplitData = oCreditHeaderData.Link
            oCreditSplitData.MoveFirst
   
            Set oInvSplitData = oInvHeaderData.Link
            oInvSplitData.MoveFirst
   
            For i = 1 To oInvSplitData.Count
               If Val(oInvSplitData.Fields.Item("NET_AMOUNT").Value) - _
                     Val(oInvSplitData.Fields.Item("AMOUNT_PAID").Value) >= _
                     Val(adoRstRpt.Fields.Item("RAmt").Value) Then
                  Exit For
               End If
               oInvSplitData.MoveNext
            Next i
            If i <= oInvSplitData.Count Then
               Set oTransactionPost = oWS.CreateObject("TransactionPost")
               bAllocated = oTransactionPost.AllocatePayment(CInt(oInvSplitData.RecordNumber), _
                              CInt(oCreditSplitData.RecordNumber), CDbl(adoRstRpt.Fields.Item("RAmt").Value), _
                              CDate(adoRstRpt.Fields.Item("RDate").Value))
               If bAllocated Then
                  laTransID(iCount) = adoRstRpt.Fields.Item("TransactionID").Value
                  iCount = iCount + 1
               End If
            End If
            Set oCreditSplitData = Nothing
            Set oInvSplitData = Nothing
         End If
      End If
      adoRstRpt.MoveNext
   Wend

   MsgBox iCount & " Transactions (out of " & iTotalTran & ") Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
   If iTotalTran - iCount > 0 Then
'     Report of posting exceptions to be printed
'      ShowReport App.Path & szReportPath & "\ReceiptExceptionReport.rpt"
   End If
   UpdateMarked laTransID, iCount, adoConn, "tlbReceipt"

   If adoPoA.EOF Then
      MsgBox "There is no Payment on Account to update to SAGE.", vbInformation + vbOKOnly, "Payment on Account "
      adoRstRpt.Close
      Set adoRstRpt = Nothing
      GoTo Updated
   End If

PoA:
'  Update Payment On Account
   iCount = 0
   ReDim laTransID(adoPoA.RecordCount) As Long

   While Not adoPoA.EOF
'      Create an instance of TransactionPost for the Sales Receipt
      Set oTransactionPost = oWS.CreateObject("TransactionPost")

'      Fill in the Header fields
      oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoPoA!SageAccountNumber)
      oTransactionPost.Header("DATE").Value = CDate(adoPoA!TDate)
      oTransactionPost.Header("POSTED_DATE").Value = CDate(Date)
      oTransactionPost.Header("TYPE").Value = CByte(sdoSA)
      oTransactionPost.Header("DETAILS").Value = "Payment on Account"
      oTransactionPost.Header("BANK_CODE").Value = CStr(adoPoA!BankCode)
      oTransactionPost.Header("Inv_Ref").Value = IIf(IsNull(adoPoA!ExtRef), "", adoPoA!ExtRef)

'      Create a split item by adding an empty split to the Items
'      collection of the TransactionPost Object.
      Set oSplitData = oTransactionPost.Items.Add()

'      Fill in the Split fields - note a Sales Receipt only has one split
      oSplitData.Fields.Item("TYPE").Value = CByte(sdoSR)
'oSplitData.Fields.Item("DEPT_NUMBER").Value = CStr(???)                   PoA does not have any department number
      oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoPoA!BankCode)
      oSplitData.Fields.Item("TAX_CODE").Value = CInt(9)
      oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoPoA!OSAllocation)
      oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(0)
      oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
      oSplitData.Fields.Item("DETAILS").Value = CStr(oTransactionPost.Header("DETAILS").Value)
      oSplitData.Fields.Item("INTERNAL_REF").Value = CStr(adoPoA.Fields("Ref").Value)

      If oTransactionPost.Update Then
         SavePostingNumberPoA "tblPoA", "TransactionID", CLng(adoPoA.Fields("TransactionID").Value), oTransactionPost.PostingNumber, adoConn
         laTransID(iCount) = adoPoA!TransactionID
         iCount = iCount + 1
      End If

      adoPoA.MoveNext
   Wend

   MsgBox "Total " & iCount & " Payment on Account transactions have been updated.", vbInformation + vbOKOnly, "Posting to SAGE"
   UpdateMarked laTransID, iCount, adoConn, "tblPoA"

Updated:

   oWS.Disconnect

'   Destroy the Objects
   Set oCreditHeaderData = Nothing
   Set oInvHeaderData = Nothing
   Set oTransactionPost = Nothing
   Set oSplitData = Nothing
   Set oHeaderData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Set adoRstRptCh = Nothing
   Set adoRstRpt = Nothing
   Set adoConn = Nothing

   Exit Sub

' Error Handling Code
Error_Handler:

   MsgBox "The SDO generated the following error: " & oSDO.LastError.text & ERR.Number & " -(pcm_SR_Posting) " & ERR.description, vbOKOnly, "Posted to SAGE"

   Set oCreditHeaderData = Nothing
   Set oInvHeaderData = Nothing
   Set oTransactionPost = Nothing
   Set oSplitData = Nothing
   Set oHeaderData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   adoRstRpt.Close
   Set adoRstRpt = Nothing
   Set adoConn = Nothing
End Sub

Private Function SagePostingNumberInv(szTransactionID As String, adoConn As ADODB.Connection) As Long
   Dim szStr As String
   Dim adoRst As New ADODB.Recordset

   szStr = "SELECT DemandRecords.SAGEPostingNumber " & _
           "FROM DemandRecords, DemandSplitRecords, tlbReceipt " & _
           "WHERE tlbReceipt.TransactionID = " & CLng(szTransactionID) & " AND " & _
               "tlbReceipt.DemandRef = DemandSplitRecords.DSR AND " & _
               "DemandSplitRecords.DemandID = DemandRecords.DemandID;"
'Debug.Print szStr
   adoRst.Open szStr, adoConn, adOpenStatic, adLockReadOnly

   SagePostingNumberInv = adoRst.Fields.Item("SAGEPostingNumber").Value

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function SagePostingNumberPoA(szTransactionID As String, adoConn As ADODB.Connection) As Long
   Dim szStr As String
   Dim adoRst As New ADODB.Recordset

   szStr = "SELECT spare2 FROM tblPoA " & _
           "WHERE TransactionID = " & CLng(szTransactionID) & ";"

   adoRst.Open szStr, adoConn, adOpenStatic, adLockReadOnly

   SagePostingNumberPoA = adoRst.Fields.Item("spare2").Value

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub UpdateMarked(lTransID() As Long, iCount As Integer, adoConn As ADODB.Connection, szTable As String)
   Dim i As Integer
   Dim SQLStr As String

   For i = 0 To iCount - 1
   'get the current password from the usernames table
      SQLStr = "UPDATE " & szTable & " " & _
               "SET UpDateSage = True, IsSageUpdate = False " & _
               "WHERE " & _
                  "TransactionID = " & lTransID(i) & ""
      adoConn.Execute SQLStr
   Next i
End Sub

Private Sub ExportSI2Sage()
   Dim iCount As Integer, iKount As Integer
   Dim szaDemandID() As String

'  Export data to sage from demand and demandsplit table
'  Error Handler
   On Error GoTo Error_Handler

'    Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oTransactionPost As SageDataObject120.TransactionPost
   Dim oSplitData As SageDataObject120.SplitData

'   Declare Variables for Database connectivity
   Dim adoConn As ADODB.Connection
   Dim adoRstDmdHd As ADODB.Recordset
   Dim adoRstDmdSpt As ADODB.Recordset

'   Declare Variables
   Dim szDataPath As String, szSQLStr As String
'
'   Set the connection to the databases
   Set adoConn = New ADODB.Connection
   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

'   Connect to Demands table to add new demands.
   Set adoRstDmdHd = New ADODB.Recordset
   szSQLStr = "SELECT DemandRecords.DemandID as D_ID, " & _
                  "DemandRecords.SageAccountNumber as S_AC, " & _
                  "DemandRecords.TransactionType as T_TYPE, " & _
                  "DemandRecords.IssueDate as I_DATE, " & _
                  "SUM(DemandSplitRecords.Amount) as AMT, " & _
                  "SUM(DemandSplitRecords.TotalAmount) as TAMT " & _
              "FROM DemandRecords, DemandSplitRecords " & _
              "WHERE DemandRecords.UPDATE_SAGE = False AND " & _
                  "DemandRecords.DEMANDID =  DemandSplitRecords.DEMANDID " & _
              "GROUP BY DemandRecords.DemandID, " & _
                  "DemandRecords.SageAccountNumber, " & _
                  "DemandRecords.TransactionType, " & _
                  "DemandRecords.IssueDate " & _
              "ORDER BY DEMANDRECORDS.DEMANDID;"

   adoRstDmdHd.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   if there no data to be exported then exit from this method
   If adoRstDmdHd.EOF Then
      MsgBox "There no demand to be exported to SAGE.", vbOKOnly, "SAGE"
      chkDemands.Value = False
      Set adoConn = Nothing
      Set adoRstDmdHd = Nothing
      Exit Sub
   End If

   Set adoRstDmdSpt = New ADODB.Recordset

'   Create the SDO Engine Object
   Set oSDO = New SageDataObject120.SDOEngine

'    Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Prestige")
'
'    Select Company.  The SelectCompany method takes the program install
'    folder as a parameter
   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
'       Select Company. The SelectCompany method takes the program install
'       folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   End If
'      Try to Connect - Will throw an exception if it fails
   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then

'        Populate Header fields

      iCount = 0
      iKount = 0
'Note:
'        The ACCOUNT_REF field must be populated with a valid
'        Customer Account Reference
      While Not adoRstDmdHd.EOF
'           Create Instances of Objects
         Set oTransactionPost = oWS.CreateObject("TransactionPost")
         szSQLStr = "SELECT DemandSplitRecords.SplitID as S_ID, " & _
                        "DemandSplitRecords.NominalCodeforAmount as NCA, " & _
                        "DemandSplitRecords.Amount AS AMT, " & _
                        "DemandSplitRecords.VATAmount AS VAMT, " & _
                        "DemandSplitRecords.SageRef AS SAGEREF, " & _
                        "DemandSplitRecords.DueDate AS D_DT, " & _
                        "DemandSplitRecords.DateFrom AS F_DT, " & _
                        "DemandSplitRecords.DateTo AS T_DT, " & _
                        "DemandSplitRecords.Description AS DESCP, " & _
                        "DemandSplitRecords.VAT_CODE AS V_CODE, " & _
                        "DemandSplitRecords.SageDepartment AS S_DEPT " & _
                    "FROM DemandSplitRecords " & _
                    "WHERE DemandSplitRecords.DemandID = " & adoRstDmdHd.Fields("D_ID").Value & " " & _
                        "AND DemandSplitRecords.DEMANDSTATEMENT=TRUE;"

         adoRstDmdSpt.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
   'Debug.Print szSQLStr
         oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstDmdHd.Fields("S_AC").Value)
         oTransactionPost.Header("BANK_CODE").Value = CStr(GlbBankID(adoRstDmdHd.Fields("D_ID").Value))

         If optIssueDt.Value Then
            oTransactionPost.Header("Date").Value = CDate(adoRstDmdHd.Fields("I_DATE").Value)                 'PLS SEE THE DOC IN SUPEROFFICE
            oTransactionPost.Header("Date_Due").Value = CDate(adoRstDmdHd.Fields("I_DATE").Value)
         End If

         If optDueDt.Value Then
            oTransactionPost.Header("Date").Value = CDate(adoRstDmdSpt.Fields("D_DT").Value)                 'PLS SEE THE DOC IN SUPEROFFICE
            oTransactionPost.Header("Date_Due").Value = CDate(adoRstDmdSpt.Fields("D_DT").Value)
         End If

         If optPostingDt.Value Then
            oTransactionPost.Header("Date").Value = CDate(Date)                 'PLS SEE THE DOC IN SUPEROFFICE
            oTransactionPost.Header("Date_Due").Value = CDate(Date)
         End If
         oTransactionPost.Header("DETAILS").Value = CStr(Left(adoRstDmdSpt.Fields("DESCP").Value, 42) & " " & Format(adoRstDmdSpt.Fields("F_DT").Value, "DD/MM/YY") & "-" & Format(adoRstDmdSpt.Fields("T_DT").Value, "DD/MM/YY"))
         oTransactionPost.Header("EURO_GROSS").Value = CDbl(adoRstDmdHd.Fields("TAMT").Value)
         oTransactionPost.Header("EURO_RATE").Value = CDbl(1)
         oTransactionPost.Header("FOREIGN_GROSS").Value = CDbl(adoRstDmdHd.Fields("TAMT").Value)
         oTransactionPost.Header("FOREIGN_RATE").Value = CDbl(1)
         oTransactionPost.Header("INTEREST_RATE").Value = CDbl(0)
         oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstDmdHd.Fields("D_ID").Value)
         oTransactionPost.Header("NET_AMOUNT").Value = CDbl(adoRstDmdHd.Fields("AMT").Value)
         oTransactionPost.Header("TYPE").Value = CByte(adoRstDmdHd.Fields("T_TYPE").Value)
         oTransactionPost.Header("Posted_Date").Value = CDate(Date)

'           Loop for the Number of Splits
'Note:
'           The Transaction can have 1 or many splits
         While Not adoRstDmdSpt.EOF
            If adoRstDmdSpt.Fields("AMT").Value > 0 Then
'                  Add a split to the Header's Item collection
               Set oSplitData = oTransactionPost.Items.Add

'                  Populate Split Fields
               oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
               oSplitData.Fields.Item("DEPT_NUMBER").Value = CInt(adoRstDmdSpt.Fields("S_DEPT").Value)
               oSplitData.Fields.Item("DETAILS").Value = CStr(Left(adoRstDmdSpt.Fields("DESCP").Value, 42) & " " & Format(adoRstDmdSpt.Fields("F_DT").Value, "DD/MM/YY") & "-" & Format(adoRstDmdSpt.Fields("T_DT").Value, "DD/MM/YY")) 'data inserting: 'Service charge description'
               oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstDmdSpt.Fields("AMT").Value) 'data inserting: 1000
               oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstDmdSpt.Fields("NCA").Value)  'data inserting: 800002
               oSplitData.Fields.Item("POSTED_DATE").Value = CDate(oTransactionPost.Header("POSTED_DATE").Value)
               oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(adoRstDmdSpt.Fields("VAMT").Value)   'data inserting: 175
               oSplitData.Fields.Item("TAX_CODE").Value = CInt(adoRstDmdSpt.Fields("V_CODE").Value)   'data inserting: 1
               oSplitData.Fields.Item("TYPE").Value = CByte(oTransactionPost.Header("TYPE").Value)
               oSplitData.Fields.Item("INTERNAL_REF").Value = CStr(adoRstDmdSpt.Fields("SageRef").Value)
            End If
            adoRstDmdSpt.MoveNext
         Wend
         adoRstDmdSpt.Close
         If oTransactionPost.Update Then
            SavePostingNumber "DemandRecords", "DemandID", CLng(adoRstDmdHd.Fields("D_ID").Value), oTransactionPost.PostingNumber, adoConn
            iCount = iCount + 1
            ReDim Preserve szaDemandID(iCount) As String
            szaDemandID(iCount - 1) = CStr(adoRstDmdHd.Fields("D_ID").Value)
         Else
            iKount = iKount + 1
         End If
         adoRstDmdHd.MoveNext
         Set oSplitData = Nothing
         Set oTransactionPost = Nothing
      Wend
      adoRstDmdHd.Close

      Dim i As Integer
      For i = 0 To iCount - 1
         ' Update the record in header table, UPDATE_SAGE = True
         adoRstDmdHd.Open "UPDATE DemandRecords " & _
                          "SET UPDATE_SAGE = TRUE " & _
                          "WHERE UPDATE_SAGE = FALSE AND " & _
                                "DemandRecords.DemandID = " & CLng(szaDemandID(i)) & "", adoConn, adOpenStatic, adLockReadOnly
      Next i
      MsgBox iCount & " Transactions (out of " & iCount + iKount & ") Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
      oWS.Disconnect
   End If

'   oWS.Disconnect
   Set oWS = Nothing
   Set oSDO = Nothing

   Set adoRstDmdSpt = Nothing
   Set adoRstDmdHd = Nothing
   Set adoConn = Nothing

   Exit Sub

'    Error Handling Code
Error_Handler:
   
   MsgBox "The SDO generated the following error: " & oSDO.LastError.text, vbOKOnly, "Posted to SAGE"
   ' Destroy Objects
   Set oTransactionPost = Nothing
   Set oSplitData = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Set adoConn = Nothing
   Set adoRstDmdHd = Nothing
End Sub

Private Sub SavePostingNumberPoA(tlbName As String, tblFieldName As String, lDemandID, lSPN As Long, adoConn As ADODB.Connection)
   Dim SQLStr As String

   'get the current password from the usernames table
   If VarType(lDemandID) = vbString Then
      SQLStr = "UPDATE " & tlbName & " " & _
               "SET spare2 = " & lSPN & " " & _
               "WHERE UpDateSage = FALSE AND " & _
                     "" & tblFieldName & " = '" & lDemandID & "';"
   Else
      SQLStr = "UPDATE " & tlbName & " " & _
               "SET spare2 = " & lSPN & " " & _
               "WHERE UpDateSage = FALSE AND " & _
                     "" & tblFieldName & " = " & lDemandID & ";"
   End If
'Debug.Print SQLStr
   adoConn.Execute SQLStr
End Sub

Private Sub SavePostingNumber(tlbName As String, tblFieldName As String, lDemandID, lSPN As Long, adoConn As ADODB.Connection)
   Dim SQLStr As String

   'get the current password from the usernames table
   If VarType(lDemandID) = vbString Then
      SQLStr = "UPDATE " & tlbName & " " & _
               "SET SAGEPostingNumber = " & lSPN & " " & _
               "WHERE UPDATE_SAGE = FALSE AND " & _
                     "" & tblFieldName & " = '" & lDemandID & "';"
   Else
      SQLStr = "UPDATE " & tlbName & " " & _
               "SET SAGEPostingNumber = " & lSPN & " " & _
               "WHERE UPDATE_SAGE = FALSE AND " & _
                     "" & tblFieldName & " = " & lDemandID & ";"
   End If
'Debug.Print SQLStr
   adoConn.Execute SQLStr
End Sub
'Private Sub SavePostingNumber(tlbName As String, tblFieldName As String, lDemandID, lSPN As Long)
'   Dim Conn As New RDO.rdoConnection
'   Dim Rst As rdoResultset
'   Dim SQLStr As String
'   'connect to the database
'   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""
'   Conn.CursorDriver = rdUseIfNeeded
'   Conn.EstablishConnection rdDriverNoPrompt
'
'   'get the current password from the usernames table
'   If VarType(lDemandID) = vbString Then
'      SQLStr = "UPDATE " & tlbName & " " & _
'               "SET SAGEPostingNumber = " & lSPN & " " & _
'               "WHERE UPDATE_SAGE = FALSE AND " & _
'                     "" & tblFieldName & " = '" & lDemandID & "';"
'   Else
'      SQLStr = "UPDATE " & tlbName & " " & _
'               "SET SAGEPostingNumber = " & lSPN & " " & _
'               "WHERE UPDATE_SAGE = FALSE AND " & _
'                     "" & tblFieldName & " = " & lDemandID & ";"
'   End If
''Debug.Print SQLStr
'   Set Rst = Conn.OpenResultset(SQLStr, rdOpenDynamic, rdConcurRowVer)
'
'   Rst.Close
'   Conn.Close
'
'   Set Rst = Nothing
'   Set Conn = Nothing
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Form_Unload 0
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   frmMMain.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   frmMMain.MousePointer = vbDefault
End Sub

Private Sub optExport_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If chkDemands.Value = 0 And chkReceipt.Value = 0 And chkPI.Value = 0 And chkBPR.Value = 0 Then
      MsgBox "Please select atleast one from the posting menu.", vbCritical + vbOKOnly, "Posting to SAGE"
      optExport.Value = False
      Exit Sub
   End If

   If chkDemands.Value = 1 Then
      If Not optIssueDt.Value And Not optDueDt.Value And Not optPostingDt.Value Then
         MsgBox "Please select a SAGE transaction date option.", vbInformation + vbOKOnly, "SAGE transaction date"
         optExport.Value = False
         Exit Sub
      End If
   End If

   fraBackupQues.Left = Frame3.Left
   fraBackupQues.Top = Image1.Top + Image1.Height + 105
   fraBackupQues.Visible = True
End Sub

Private Sub chkDemands_Click()
   optPreview.Value = True
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub
