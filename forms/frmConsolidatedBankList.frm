VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsolidatedBankList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidated Bank List"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   Icon            =   "frmConsolidatedBankList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10905
   Begin VB.TextBox txtBankCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1845
      MaxLength       =   10
      TabIndex        =   0
      Top             =   135
      Width           =   2625
   End
   Begin VB.TextBox txtBankACNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      MaxLength       =   8
      TabIndex        =   2
      Top             =   495
      Width           =   2620
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   1
      Top             =   135
      Width           =   2625
   End
   Begin VB.TextBox txtBANK_ID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   11
      Top             =   945
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5805
      TabIndex        =   8
      Top             =   4770
      Width           =   1260
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   1710
      TabIndex        =   5
      Top             =   4770
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   4770
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7155
      TabIndex        =   9
      Top             =   4770
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4500
      TabIndex        =   7
      Top             =   4770
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3060
      TabIndex        =   6
      Top             =   4770
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxConsolidateBank 
      Height          =   3195
      Left            =   75
      TabIndex        =   4
      Top             =   1515
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   5636
      _Version        =   393216
      ForeColor       =   0
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox cboSortCode 
      Height          =   300
      Left            =   1845
      TabIndex        =   3
      Top             =   495
      Width           =   2400
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4233;529"
      BoundColumn     =   2
      ColumnCount     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Code"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   135
      TabIndex        =   21
      Top             =   1290
      Width           =   2280
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Code:"
      Height          =   195
      Left            =   180
      TabIndex        =   20
      Top             =   90
      Width           =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Code"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   4860
      TabIndex        =   19
      Top             =   1305
      Width           =   3135
   End
   Begin VB.Label lblBank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   7740
      TabIndex        =   18
      Top             =   1305
      Width           =   2190
   End
   Begin VB.Label lblBankID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2205
      TabIndex        =   17
      Top             =   1290
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Code:"
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   540
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank ID:"
      Height          =   195
      Index           =   0
      Left            =   6255
      TabIndex        =   15
      Top             =   990
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account Name:"
      Height          =   195
      Left            =   4725
      TabIndex        =   14
      Top             =   180
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account Number:"
      Height          =   195
      Left            =   4725
      TabIndex        =   13
      Top             =   540
      Width           =   1665
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   75
      TabIndex        =   12
      Top             =   1275
      Width           =   10545
   End
End
Attribute VB_Name = "frmConsolidatedBankList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BANK_NEW_ENTRY_ As Boolean
'Dim GRID_BROWSE_ As Boolean
Public CALLER_FORM As String
''PLease put a unique key validation on bank code field othewise on update query of closing balance field from bank reconciliation will update this bal incoorecttlyy
Private Const IDC_APPSTARTING = 32650&
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Const IDC_CROSS = 32515&
Private Const IDC_IBEAM = 32513&
Private Const IDC_ICON = 32641&
Private Const IDC_NO = 32648&
Private Const IDC_SIZE = 32640&
Private Const IDC_SIZEALL = 32646&
Private Const IDC_SIZENESW = 32643&
Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_UPARROW = 32516&
Private Const IDC_WAIT = 32514&
Private Declare Function LoadCursorLong Lib "User32" Alias "LoadCursorA" _
  (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Declare Function SetCursor Lib "User32" _
  (ByVal hCursor As Long) As Long


Private Sub cboSortCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtBankACNumber
    End If
End Sub

Private Sub cmdCancel_Click()
   ComponentEnableMode Me, NewEntryMode, flxConsolidateBank
   ComponentEnableMode Me, DefaultMode, flxConsolidateBank
   'cboSortCode.text = ""
   cboSortCode.ListIndex = -1
   txtBANK_ID.Enabled = True
End Sub

Private Sub cmdClose_Click()
     Unload Me
End Sub



Private Sub cmdDelete_Click()
    If txtBANK_ID.text = "" Then
        MsgBox "Please select a consolidated bank to delete"
        Exit Sub
    Else
        Dim adoConn As New ADODB.Connection
        If MsgBox("Are you sure you want to delete selectd Consolidated Bank?", vbYesNo, "Please Confirm") = vbYes Then
            adoConn.Open getConnectionString
            adoConn.Execute "Delete from ConsolidatedBankList where conBankID=" & txtBANK_ID.text & ""
            adoConn.Close
            MsgBox "Selected Consolidated Bank information has been deleted", vbInformation, "Record Deleted"
            Call loadflxConsolidateBank
            Set adoConn = Nothing
         End If
    End If
End Sub

Private Sub cmdEdit_Click()
   If flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 1) = "" Then
       MsgBox "Please select a record to continue.", vbInformation, "Edit record"
       Exit Sub
   End If
   FocusControl txtBankCode
   BANK_NEW_ENTRY_ = False
   ComponentEnableMode Me, EditMode, flxConsolidateBank
   txtBANK_ID.Enabled = False
End Sub

Private Sub cmdNew_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   adoConn.Open getConnectionString
   BANK_NEW_ENTRY_ = True
   ComponentEnableMode Me, NewEntryMode, flxConsolidateBank
   FocusControl txtBankName
   txtBANK_ID.text = NextRef(adoConn, "BANK_ID")
   adoConn.Close
   Set adoConn = Nothing
   FocusControl txtBankCode
End Sub
Private Function NextID(adoConn As ADODB.Connection) As Long
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   szSQL = "SELECT MAX(Cint(conBankID))+1 AS Ref FROM ConsolidatedBankList;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        NextID = IIf(adoRst.EOF, 1, IIf(IsNull(adoRst!ref), 1, adoRst!ref))
   adoRst.Close
   Set adoRst = Nothing
End Function
Private Sub UpdateRecord(adoConn As ADODB.Connection)
        Dim rsBank As New ADODB.Recordset
        If txtBANK_ID.text = "" Then
            MsgBox "Bank ID not found", vbInformation, "Warning"
            Exit Sub
        End If
        
        If txtBankName.text = "" Then
            MsgBox "Please enter Bank Name ", vbInformation, "Warning"
            FocusControl txtBankName
            Exit Sub
        End If
        If cboSortCode.text = "" Then
            ShowMsgInTaskBar "Please enter the sort code", "Y", "N"
            cboSortCode.SetFocus
            Exit Sub
        End If
        rsBank.Open "Select * from ConsolidatedBankList where conBankID=" & txtBANK_ID.text & "", adoConn, adOpenDynamic, adLockPessimistic
        If rsBank.EOF Then
            rsBank.Close
            Set rsBank = Nothing
            Exit Sub
        Else
            rsBank!conBankID = txtBANK_ID.text
            rsBank!BankName = txtBankName.text
            rsBank!BankCode = txtBankCode.text
            rsBank!BankACNumber = txtBankACNumber.text
            rsBank!SortCode = cboSortCode.text
            rsBank.Update
        End If
    
End Sub
Private Sub cmdSave_Click()
    
   If txtBankName.text = "" Then
      ShowMsgInTaskBar "Please enter the bank Name", "Y", "N"
      txtBankName.SetFocus
      Exit Sub
   End If
   If txtBankCode.text = "" Then
      ShowMsgInTaskBar "Please enter the bank Code", "Y", "N"
      txtBankCode.SetFocus
      Exit Sub
   End If
   
   If txtBankACNumber.text = "" Then
      ShowMsgInTaskBar "Please enter the bank AC number", "Y", "N"
      txtBankACNumber.SetFocus
      Exit Sub
   End If
   If Len(txtBankACNumber.text) <> 8 Then
        MsgBox "Consolidated Bank account number must be 8 character long", vbOKOnly, "Warning"
        FocusControl txtBankACNumber
        Exit Sub
   End If
   If cboSortCode.text = "" Then
      ShowMsgInTaskBar "Please enter the sort code", "Y", "N"
      cboSortCode.SetFocus
      Exit Sub
   End If
    If Len(cboSortCode.text) <> 6 Then
        MsgBox "Consolidated Bank sort code must be 6 character long", vbOKOnly, "Warning"
        FocusControl cboSortCode
        Exit Sub
   End If
   Dim rsConsBank As New ADODB.Recordset
   Dim sSQLQuery_, sFilter As String
   Dim adoConn As New ADODB.Connection
  
   adoConn.Open getConnectionString
   Dim rsConsolidatedBank As New ADODB.Recordset
   If BANK_NEW_ENTRY_ Then
        sSQLQuery_ = "SELECT * " & _
                       "FROM ConsolidatedBankList where BankCode='" & txtBankCode.text & "'"
                       rsConsolidatedBank.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic
                       
        
        If rsConsolidatedBank.State = 1 Then
            rsConsolidatedBank.Close
        End If
        rsConsolidatedBank.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
        
        If Not rsConsolidatedBank.EOF Then
                 MsgBox "This bank code already exists", vbInformation, "Warning"
                 rsConsolidatedBank.Close
                 adoConn.Close
                 Set adoConn = Nothing
                 Exit Sub
        End If
        rsConsolidatedBank.Close
        Set rsConsolidatedBank = Nothing
    End If
   'sFilter = "WHERE ConsolidatedBankList.BANK_ID = " & txtBANK_ID.text & " "

   If BANK_NEW_ENTRY_ Then
      sSQLQuery_ = "SELECT * " & _
                   "FROM ConsolidatedBankList where 1=2"
       rsConsolidatedBank.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic
   Else
'      sSQLQuery_ = "SELECT * " & _
'                   "FROM ConsolidatedBankList " & sFilter
        If BANK_NEW_ENTRY_ = False Then
              sSQLQuery_ = "SELECT * " & _
                              "FROM ConsolidatedBankList where BankCode='" & txtBankCode.text & "' AND conBankID <> " & txtBANK_ID.text & ""
                              'rsConsolidatedBank.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic
                              
               rsConsBank.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
               
               If Not rsConsBank.EOF Then
                        MsgBox "This bank code already exists", vbInformation, "Warning"
                        rsConsBank.Close
                        adoConn.Close
                        Set adoConn = Nothing
                        Exit Sub
               End If
               rsConsBank.Close
               Set rsConsBank = Nothing
               
                sSQLQuery_ = "SELECT * " & _
                              "FROM tlbBankReconClosingBal where ClientID='" & flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 2) & "'"
                              'rsConsBank.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic
                              
               rsConsBank.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
               
               If Not rsConsBank.EOF And flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 2) <> txtBankCode.text Then
                        MsgBox "You cannot change the Bank Code as some transcation exists under this Bank Code", vbInformation, "Warning"
                        rsConsBank.Close
                        adoConn.Close
                        Set adoConn = Nothing
                        cmdCancel_Click
                        Exit Sub
               End If
               rsConsBank.Close
               Set rsConsBank = Nothing
               
         End If
      adoConn.Execute "UPDATE tlbClientBanks " & _
                      "SET ConsBankACNumber = '" & txtBankACNumber.text & "',ConsSortCode = '" & cboSortCode.text & "',conBankCode='" & txtBankCode.text & "' " & _
                      "WHERE ConsolidatedBANKID = " & txtBANK_ID.text & ";"
   End If

   

   If BANK_NEW_ENTRY_ Then
        rsConsolidatedBank.AddNew
        txtBANK_ID.text = NextID(adoConn)
        rsConsolidatedBank!conBankID = txtBANK_ID.text
        rsConsolidatedBank!BankName = txtBankName.text
        rsConsolidatedBank!BankCode = txtBankCode.text
        rsConsolidatedBank!BankACNumber = txtBankACNumber.text
        rsConsolidatedBank!SortCode = cboSortCode.text
        rsConsolidatedBank.Update
        rsConsolidatedBank.Close
   Else
        Call UpdateRecord(adoConn)
   End If


'   AddNewRef adoConn, "BANK_ID", CLng(txtBANK_ID.text)
   adoConn.Close
   Set adoConn = Nothing

   ComponentEnableMode Me, DefaultMode, flxConsolidateBank
   configflxConsolidateBank
   loadflxConsolidateBank
   txtBANK_ID.Enabled = True
   MsgBox "Consolidated Bank Information has benn saved successfully", vbInformation, "Saved!"
   FocusControl cmdClose
End Sub

Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   ComponentEnableMode Me, DefaultMode, flxConsolidateBank

   configflxConsolidateBank
   loadflxConsolidateBank
   LoadCboSortCode
End Sub
Private Sub LoadCboSortCode()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rsSortCode As New ADODB.Recordset
    rsSortCode.Open "Select Distinct sort_Code from tlbBank", adoConn, adOpenStatic, adLockReadOnly
    cboSortCode.Clear
    While Not rsSortCode.EOF
        cboSortCode.AddItem rsSortCode("sort_Code").Value
        rsSortCode.MoveNext
    Wend
    rsSortCode.Close
    Set rsSortCode = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Function ComponentEnableMode(ByVal frmCurrent As Form, ByVal mode As ComponentMode, ByVal flxConsolidateBank As MSHFlexGrid)
    Dim ctrl As Control

    Select Case mode
        Case ComponentMode.DefaultMode, ComponentMode.GridLostFocus
            For Each ctrl In frmCurrent.Controls
                Select Case TypeName(ctrl)
                    Case "TextBox"
                        ctrl.Locked = True
                        ctrl.text = ""
                    Case "CheckBox", "DataCombo"
                        ctrl.Enabled = False
                End Select
            Next ctrl
            flxConsolidateBank.Enabled = True
            cboSortCode.Locked = True

            frmCurrent.cmdNew.Enabled = True
            frmCurrent.cmdEdit.Enabled = False
            frmCurrent.cmdSave.Enabled = False
            frmCurrent.cmdCancel.Enabled = False
            frmCurrent.cmdClose.Enabled = True
            frmCurrent.cmdDelete.Enabled = True

        Case ComponentMode.GridRowOnSelection
            For Each ctrl In frmCurrent.Controls
                Select Case TypeName(ctrl)
                    Case "TextBox"
                       ' ctrl.Enabled = False
                        ctrl.text = ""
                    Case "CheckBox", "DataCombo"
                        ctrl.Enabled = True
                End Select
            Next ctrl
            flxConsolidateBank.Enabled = True

            frmCurrent.cmdNew.Enabled = True
            frmCurrent.cmdEdit.Enabled = True
            frmCurrent.cmdSave.Enabled = False
            frmCurrent.cmdCancel.Enabled = False
            frmCurrent.cmdClose.Enabled = True
            frmCurrent.cmdDelete.Enabled = True

        Case ComponentMode.NewEntryMode
            For Each ctrl In frmCurrent.Controls
                Select Case TypeName(ctrl)
                    Case "TextBox"
                        ctrl.Enabled = True
                        ctrl.Locked = False
                        ctrl.text = ""
                    Case "DataCombo"
                        ctrl.Enabled = True
                End Select
            Next ctrl
            flxConsolidateBank.Enabled = False
            cboSortCode.Locked = False
            frmCurrent.cmdNew.Enabled = False
            frmCurrent.cmdEdit.Enabled = False
            frmCurrent.cmdSave.Enabled = True
            frmCurrent.cmdCancel.Enabled = True
            frmCurrent.cmdClose.Enabled = False
            frmCurrent.cmdDelete.Enabled = False

        Case ComponentMode.EditMode
            For Each ctrl In frmCurrent.Controls
                Select Case TypeName(ctrl)
                    Case "TextBox", "CheckBox", "DataCombo", "ComboBox"
                        ctrl.Enabled = True
                         ctrl.Locked = False
                End Select
            Next ctrl
            flxConsolidateBank.Enabled = False
             cboSortCode.Locked = False
            frmCurrent.cmdNew.Enabled = False
            frmCurrent.cmdEdit.Enabled = False
            frmCurrent.cmdSave.Enabled = True
            frmCurrent.cmdCancel.Enabled = True
            frmCurrent.cmdClose.Enabled = False
            frmCurrent.cmdDelete.Enabled = False

        Case ComponentMode.GridLostFocus
            For Each ctrl In frmCurrent.Controls
                Select Case TypeName(ctrl)
                    Case "TextBox", "CheckBox", "DataCombo"
                        'ctrl.Enabled = False
                End Select
            Next ctrl
            flxConsolidateBank.Enabled = True

            frmCurrent.cmdNew.Enabled = True
            frmCurrent.cmdEdit.Enabled = False
            frmCurrent.cmdSave.Enabled = False
            frmCurrent.cmdCancel.Enabled = False
            frmCurrent.cmdClose.Enabled = True
            frmCurrent.cmdDelete.Enabled = True
    End Select
End Function
Private Sub configflxConsolidateBank()
   Dim szHeader As String

   flxConsolidateBank.Cols = 6
   flxConsolidateBank.Clear
   szHeader$ = "|<ID|<Name|<Type"
   flxConsolidateBank.FormatString = szHeader$
   flxConsolidateBank.ColWidth(0) = 15        'Solid column
   flxConsolidateBank.ColWidth(1) = 0        'Client ID
   flxConsolidateBank.ColWidth(2) = 2775       'Bank name
   flxConsolidateBank.ColWidth(3) = 2280       'Bank Code
   flxConsolidateBank.ColWidth(4) = 2190        'Account number
   flxConsolidateBank.ColWidth(5) = 3135        'Sort code
   flxConsolidateBank.Rows = 2
   flxConsolidateBank.RowHeight(0) = 0
   
   
End Sub
Public Function loadflxConsolidateBank()
   Dim sSQLQuery_ As String, szHeader As String
   Dim Conn As New ADODB.Connection
   Dim rsConsolidatedBank As New ADODB.Recordset
   Conn.Open getConnectionString
   sSQLQuery_ = "Select conBankID,BankName,BankACNumber,SortCode,BankCode from ConsolidatedBankList order by conBankID"
   rsConsolidatedBank.Open sSQLQuery_, Conn, adOpenStatic, adLockReadOnly
   Dim iRow  As Integer
   iRow = 1
   flxConsolidateBank.Rows = rsConsolidatedBank.RecordCount + 1
   While Not rsConsolidatedBank.EOF
        flxConsolidateBank.TextMatrix(iRow, 1) = rsConsolidatedBank!conBankID
        flxConsolidateBank.TextMatrix(iRow, 2) = IIf(IsNull(rsConsolidatedBank!BankCode), "", rsConsolidatedBank!BankCode)
        flxConsolidateBank.TextMatrix(iRow, 3) = rsConsolidatedBank!BankName
        flxConsolidateBank.TextMatrix(iRow, 4) = rsConsolidatedBank!SortCode
        flxConsolidateBank.TextMatrix(iRow, 5) = rsConsolidatedBank!BankACNumber
        iRow = iRow + 1
        rsConsolidatedBank.MoveNext
   Wend
   rsConsolidatedBank.Close
   Set rsConsolidatedBank = Nothing
   Conn.Close
   Set Conn = Nothing
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  SetMouseCursor IDC_ARROW
End Sub
Private Function SetMouseCursor(CursorType As Long)
  Dim hCursor As Long
  hCursor = LoadCursorLong(0&, CursorType)
  hCursor = SetCursor(hCursor)
End Function
Private Sub flxConsolidateBank_Click()
    On Error GoTo ERR
        If flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 1) = "" Then Exit Sub
        ComponentEnableMode Me, GridRowOnSelection, flxConsolidateBank
        txtBANK_ID.text = flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 1)
        txtBankCode.text = flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 2)
        txtBankName.text = flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 3)
        cboSortCode.text = flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 4)
        txtBankACNumber.text = flxConsolidateBank.TextMatrix(flxConsolidateBank.row, 5)
   Exit Sub
ERR:
   
End Sub


Private Sub flxConsolidateBank_LostFocus()
   'GRID_BROWSE_ = False
   txtBANK_ID.Enabled = True
End Sub

Private Sub flxConsolidateBank_RowColChange()
   'GRID_BROWSE_ = True
   populateControl Me, flxConsolidateBank
End Sub



Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
End Sub

Private Sub lblBank_Click()
'issue 0559: Bank Details not sorting correctly
'Written by anol 14 May 2015
    Dim sSQLQuery_ As String, szHeader As String
   Dim conBank As New ADODB.Connection

   conBank.Open getConnectionString
   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
                "FROM tlbBank Order by BANK_NAME "

   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
   flxConsolidateBank.Rows = 2

   populateGridDefineHeader conBank, sSQLQuery_, flxConsolidateBank, szHeader

   flxConsolidateBank.RowHeight(0) = 0
   flxConsolidateBank.ColWidth(0) = "1000"
   flxConsolidateBank.ColWidth(1) = "1200"
   flxConsolidateBank.ColWidth(2) = "2000"
   flxConsolidateBank.ColWidth(3) = "2000"
   flxConsolidateBank.ColWidth(4) = 0
   flxConsolidateBank.ColWidth(5) = 0
   flxConsolidateBank.ColWidth(6) = 0
   flxConsolidateBank.ColWidth(7) = 0
   flxConsolidateBank.ColWidth(8) = "1500"
   conBank.Close
   Set conBank = Nothing
End Sub
    

'Private Sub lblBankID_Click()
'    'issue 0559: Bank Details not sorting correctly
''Written by anol 14 May 2015
'    Dim sSQLQuery_ As String, szHeader As String
'   Dim conBank As New ADODB.Connection
'
'   conBank.Open getConnectionString
'   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
'                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
'                "FROM tlbBank Order by cint(BANK_ID) "
'
'   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
'   flxConsolidateBank.Rows = 2
'
'   populateGridDefineHeader conBank, sSQLQuery_, flxConsolidateBank, szHeader
'
'   flxConsolidateBank.RowHeight(0) = 0
'   flxConsolidateBank.ColWidth(0) = "1000"
'   flxConsolidateBank.ColWidth(1) = "1200"
'   flxConsolidateBank.ColWidth(2) = "2000"
'   flxConsolidateBank.ColWidth(3) = "2000"
'   flxConsolidateBank.ColWidth(4) = 0
'   flxConsolidateBank.ColWidth(5) = 0
'   flxConsolidateBank.ColWidth(6) = 0
'   flxConsolidateBank.ColWidth(7) = 0
'   flxConsolidateBank.ColWidth(8) = "1500"
'   conBank.Close
'   Set conBank = Nothing
'
'End Sub
'
'Private Sub lblbranch_Click()
'      'issue 0559: Bank Details not sorting correctly
''Written by anol 14 May 2015
'    Dim sSQLQuery_ As String, szHeader As String
'   Dim conBank As New ADODB.Connection
'
'   conBank.Open getConnectionString
'   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
'                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
'                "FROM tlbBank Order by BANK_BRANCH "
'
'   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
'   flxConsolidateBank.Rows = 2
'
'   populateGridDefineHeader conBank, sSQLQuery_, flxConsolidateBank, szHeader
'
'   flxConsolidateBank.RowHeight(0) = 0
'   flxConsolidateBank.ColWidth(0) = "1000"
'   flxConsolidateBank.ColWidth(1) = "1200"
'   flxConsolidateBank.ColWidth(2) = "2000"
'   flxConsolidateBank.ColWidth(3) = "2000"
'   flxConsolidateBank.ColWidth(4) = 0
'   flxConsolidateBank.ColWidth(5) = 0
'   flxConsolidateBank.ColWidth(6) = 0
'   flxConsolidateBank.ColWidth(7) = 0
'   flxConsolidateBank.ColWidth(8) = "1500"
'   conBank.Close
'   Set conBank = Nothing
'End Sub
'
'Private Sub lblSortCode_Click()
'     'issue 0559: Bank Details not sorting correctly
''Written by anol 14 May 2015
'    Dim sSQLQuery_ As String, szHeader As String
'   Dim conBank As New ADODB.Connection
'
'   conBank.Open getConnectionString
'   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
'                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
'                "FROM tlbBank Order by SORT_CODE "
'
'   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
'   flxConsolidateBank.Rows = 2
'
'   populateGridDefineHeader conBank, sSQLQuery_, flxConsolidateBank, szHeader
'
'   flxConsolidateBank.RowHeight(0) = 0
'   flxConsolidateBank.ColWidth(0) = "1000"
'   flxConsolidateBank.ColWidth(1) = "1200"
'   flxConsolidateBank.ColWidth(2) = "2000"
'   flxConsolidateBank.ColWidth(3) = "2000"
'   flxConsolidateBank.ColWidth(4) = 0
'   flxConsolidateBank.ColWidth(5) = 0
'   flxConsolidateBank.ColWidth(6) = 0
'   flxConsolidateBank.ColWidth(7) = 0
'   flxConsolidateBank.ColWidth(8) = "1500"
'   conBank.Close
'   Set conBank = Nothing
'End Sub

Private Sub txtBankACNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
End Sub




Private Sub txtBankCode_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        FocusControl txtBankName
    End If
End Sub

Private Sub txtBankName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cboSortCode
    End If
End Sub


'Private Sub cboSortCode_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
'        FocusControl cmdSave
'    End If
'End Sub
