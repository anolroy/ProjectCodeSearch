VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank List"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8160
   Begin VB.TextBox txtBANK_ADDRESS1 
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
      Left            =   5130
      MaxLength       =   70
      TabIndex        =   4
      Top             =   135
      Width           =   2620
   End
   Begin VB.TextBox txtBANK_ADDRESS2 
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
      Left            =   5130
      MaxLength       =   70
      TabIndex        =   5
      Top             =   495
      Width           =   2620
   End
   Begin VB.TextBox txtBANK_ADDRESS3 
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
      Left            =   5130
      MaxLength       =   70
      TabIndex        =   6
      Top             =   855
      Width           =   2620
   End
   Begin VB.TextBox txtBANK_ADDRESS4 
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
      Left            =   5130
      MaxLength       =   70
      TabIndex        =   7
      Top             =   1215
      Width           =   2620
   End
   Begin VB.TextBox txtSORT_CODE 
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
      Left            =   1215
      MaxLength       =   6
      TabIndex        =   1
      Top             =   585
      Width           =   2620
   End
   Begin VB.TextBox txtBANK_NAME 
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
      Left            =   1215
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1035
      Width           =   2620
   End
   Begin VB.TextBox txtBANK_BRANCH 
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
      Left            =   1215
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1485
      Width           =   2620
   End
   Begin VB.TextBox txtBANK_POST_CODE 
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
      Left            =   5130
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1575
      Width           =   2620
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
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Top             =   135
      Width           =   2620
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4185
      TabIndex        =   27
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6930
      TabIndex        =   14
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5580
      TabIndex        =   13
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1665
      TabIndex        =   11
      Top             =   5760
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   375
      Left            =   360
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Main"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMain 
      Height          =   3195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   7935
      _ExtentX        =   13996
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   8040
      Y1              =   6220
      Y2              =   6220
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   615
      Left            =   120
      Top             =   5640
      Width           =   7935
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   6360
      TabIndex        =   26
      Top             =   2055
      Width           =   735
   End
   Begin VB.Label lblbranch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2400
      TabIndex        =   25
      Top             =   2055
      Width           =   510
   End
   Begin VB.Label lblBank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   4320
      TabIndex        =   24
      Top             =   2055
      Width           =   840
   End
   Begin VB.Label lblSortCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Code"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   1200
      TabIndex        =   23
      Top             =   2055
      Width           =   705
   End
   Begin VB.Label lblBankID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank ID"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   135
      TabIndex        =   22
      Top             =   2055
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
      Height          =   195
      Left            =   4200
      TabIndex        =   15
      Top             =   1605
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   195
      Left            =   4200
      TabIndex        =   21
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Branch:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   1515
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank ID:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Code:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   750
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
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   7935
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BANK_NEW_ENTRY_ As Boolean
'Dim GRID_BROWSE_ As Boolean
Public CALLER_FORM As String
''

Private Sub cmdCancel_Click()
   ComponentEnableMode Me, NewEntryMode, gridMain
   ComponentEnableMode Me, DefaultMode, gridMain
   txtBANK_ID.Enabled = True
End Sub

Private Sub cmdClose_Click()
   'Unload Me
   Form_Unload 0
End Sub

Private Sub cmdDelete_Click()
' ***************************************************************************
' *************** To call the method: ComponentEnableMode, there should be a
' ***************         a command button cmdDelete.
' ***************************************************************************
    If txtBANK_ID.text = "" Then
        MsgBox "Please select a  bank to delete"
        Exit Sub
    Else
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        Dim rsBank As New ADODB.Recordset
        rsBank.Open "Select * from tlbClientBanks where Bank_ID='" & txtBANK_ID.text & "'", adoConn, adOpenStatic, adLockReadOnly
        
        If Not rsBank.EOF Then
                 MsgBox "This bank iformation cannot be deleted. Related information exists", vbInformation, "Related information exists"
                 rsBank.Close
                 Set rsBank = Nothing
                 adoConn.Close
                 Set adoConn = Nothing
                Exit Sub
        End If
        If MsgBox("Are you sure you want to delete selectd selected Bank?", vbYesNo, "Please Confirm") = vbYes Then
           
            adoConn.Execute "Delete from tlbBank where BankID=" & txtBANK_ID.text & ""
            adoConn.Close
            MsgBox "Selected Bank information has been deleted", vbInformation, "Record Deleted"
            Call LoadGrid
            adoConn.Close
            Set adoConn = Nothing
         End If
    End If
End Sub

Private Sub cmdEdit_Click()
   If gridMain.TextMatrix(gridMain.row, 0) = "" Then
       MsgBox "Please select a record to continue.", vbInformation, "Edit record"
       Exit Sub
   End If
   BANK_NEW_ENTRY_ = False
   ComponentEnableMode Me, EditMode, gridMain
   txtSORT_CODE.Locked = False
   txtBANK_ID.Enabled = False
End Sub

Private Sub cmdNew_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ";"

   BANK_NEW_ENTRY_ = True
   ComponentEnableMode Me, NewEntryMode, gridMain
   txtSORT_CODE.SetFocus
   txtBANK_ID.text = NextRef(adoConn, "BANK_ID")

   adoConn.Close
   Set adoConn = Nothing
End Sub
Private Sub UpdateRecord(adoConn As ADODB.Connection)
    Dim rsBank As New ADODB.Recordset
    
    If txtBANK_ID.text = "" Then
       MsgBox "Bank ID not found", vbInformation, "Warning"
       Exit Sub
    End If
    If txtSORT_CODE.text = "" Then
      ShowMsgInTaskBar "Please enter the sort code", "Y", "N"
      txtSORT_CODE.SetFocus
      Exit Sub
   End If
   If txtBANK_NAME.text = "" Then
       MsgBox "Bank ID not found", vbInformation, "Warning"
       FocusControl txtBANK_NAME
       Exit Sub
    End If
     If txtBANK_BRANCH.text = "" Then
       MsgBox "Please enter BANK BRANCH", vbInformation, "Warning"
       FocusControl txtBANK_BRANCH
       Exit Sub
    End If
    rsBank.Open "Select * from tlbBank where BANK_ID='" & txtBANK_ID.text & "'", adoConn, adOpenDynamic, adLockPessimistic
    If rsBank.EOF Then
        rsBank.Close
        Set rsBank = Nothing
        Exit Sub
    Else
        rsBank!BANK_ID = txtBANK_ID.text
        rsBank!SORT_CODE = txtSORT_CODE.text
        rsBank!BANK_BRANCH = txtBANK_BRANCH.text
        rsBank!BANK_NAME = txtBANK_NAME.text
        rsBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
        rsBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
        rsBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
        rsBank!BANK_ADDRESS4 = txtBANK_ADDRESS4.text
        rsBank!BANK_POST_CODE = txtBANK_POST_CODE.text
        rsBank.Update
    End If
    
End Sub
Private Sub cmdSave_Click()
   If txtSORT_CODE.text = "" Then
      ShowMsgInTaskBar "Please enter the sort code", "Y", "N"
      txtSORT_CODE.SetFocus
      Exit Sub
   End If

  If Len(txtSORT_CODE.text) <> 6 Then
        MsgBox "Sort code must be 6 character long", vbOKOnly, "Warning"
        FocusControl txtSORT_CODE
        Exit Sub
   End If
   
   Dim sSQLQuery_, sFilter As String
   Dim adoConn As New ADODB.Connection

   adoMain.ConnectionString = getConnectionString
   adoConn.Open getConnectionString

   sFilter = "WHERE tlbBank.BANK_ID = '" & txtBANK_ID.text & "' "

   If BANK_NEW_ENTRY_ Then
      sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, " & _
                     "BANK_ADDRESS1, BANK_ADDRESS2, BANK_ADDRESS3, " & _
                     "BANK_ADDRESS4, BANK_POST_CODE " & _
                   "FROM tlbBank " & sFilter
   Else
      sSQLQuery_ = "SELECT SORT_CODE, BANK_BRANCH, BANK_NAME, " & _
                     "BANK_ADDRESS1, BANK_ADDRESS2, BANK_ADDRESS3, " & _
                     "BANK_ADDRESS4, BANK_POST_CODE " & _
                   "FROM tlbBank " & sFilter
      
      adoConn.Execute "UPDATE tlbClientBanks " & _
                      "SET BANK_SC = '" & txtSORT_CODE.text & "' " & _
                      "WHERE BANK_ID = '" & txtBANK_ID.text & "';"
   End If

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If BANK_NEW_ENTRY_ Then
       Add_NoQuery Me, adoMain
   Else
'       Update_NoQuery Me, adoMain
        Call UpdateRecord(adoConn)
   End If


'   AddNewRef adoConn, "BANK_ID", CLng(txtBANK_ID.text)
   adoConn.Close
   Set adoConn = Nothing

   ComponentEnableMode Me, DefaultMode, gridMain
   LoadGrid
   txtBANK_ID.Enabled = True
   MsgBox "Bank Information has benn saved successfully", vbInformation, "Saved!"
End Sub

Private Sub Form_Load()
'   Me.Top = 0 ' (frmMMain.Height / 2) - (Me.Height / 2) - 400
'   Me.Left = 0 '(frmMMain.Width / 2) - (Me.Width / 2) - 400
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   ComponentEnableMode Me, DefaultMode, gridMain
   adoMain.ConnectionString = getConnectionString
   LoadGrid
   
   gridMain.ColWidth(0) = "1000"
   gridMain.ColWidth(1) = "1200"
   gridMain.ColWidth(2) = "2000"
   gridMain.ColWidth(3) = "2000"
   gridMain.ColWidth(4) = 0
   gridMain.ColWidth(5) = 0
   gridMain.ColWidth(6) = 0
   gridMain.ColWidth(7) = 0
   gridMain.ColWidth(8) = "1500"
End Sub
Private Function ComponentEnableMode(ByVal frmCurrent As Form, ByVal mode As ComponentMode, ByVal gridMain As MSHFlexGrid)
    Dim ctrl As Control

    Select Case mode
        Case ComponentMode.DefaultMode, ComponentMode.GridLostFocus
            For Each ctrl In frmCurrent.Controls
                Select Case TypeName(ctrl)
                    Case "TextBox"
                        'ctrl.Enabled = False
                        ctrl.text = ""
                    Case "CheckBox", "DataCombo"
                        ctrl.Enabled = False
                End Select
            Next ctrl
            gridMain.Enabled = True

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
            gridMain.Enabled = True

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
                        ctrl.text = ""
                    Case "DataCombo"
                        ctrl.Enabled = True
                End Select
            Next ctrl
            gridMain.Enabled = False

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
                End Select
            Next ctrl
            gridMain.Enabled = False

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
            gridMain.Enabled = True

            frmCurrent.cmdNew.Enabled = True
            frmCurrent.cmdEdit.Enabled = False
            frmCurrent.cmdSave.Enabled = False
            frmCurrent.cmdCancel.Enabled = False
            frmCurrent.cmdClose.Enabled = True
            frmCurrent.cmdDelete.Enabled = True
    End Select
End Function
Public Function LoadGrid()
  'issue 0559: Bank Details not sorting correctly
'Written by anol 14 May 2015
   Dim sSQLQuery_ As String, szHeader As String
   Dim conBank As New ADODB.Connection
'Order by Cint(BANK_ID) has been added by anol
   conBank.Open getConnectionString
   
   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
                "FROM tlbBank Order by Cint(BANK_ID) "

   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
   gridMain.Rows = 2

   populateGridDefineHeader conBank, sSQLQuery_, gridMain, szHeader

   gridMain.RowHeight(0) = 0
   gridMain.ColWidth(0) = "1000"
   gridMain.ColWidth(1) = "1200"
   gridMain.ColWidth(2) = "2000"
   gridMain.ColWidth(3) = "2000"
   gridMain.ColWidth(4) = 0
   gridMain.ColWidth(5) = 0
   gridMain.ColWidth(6) = 0
   gridMain.ColWidth(7) = 0
   gridMain.ColWidth(8) = "1500"
   conBank.Close
   Set conBank = Nothing
End Function

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   frmMMain.MousePointer = vbArrow
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   If CALLER_FORM = "frmClientNew4" Then
      frmClientNew4.PopulateBank
      frmClientNew4.RefreshedBankDetails
      frmClientNew4.Show
      CALLER_FORM = ""
   End If
   Unload Me
End Sub

Private Sub gridMain_Click()
   If gridMain.TextMatrix(1, 0) = "" Then Exit Sub
   'GRID_BROWSE_ = True
   ComponentEnableMode Me, GridRowOnSelection, gridMain
   populateControl Me, gridMain
End Sub

Private Sub gridMain_GotFocus()
   'GRID_BROWSE_ = True
End Sub

Private Sub gridMain_LostFocus()
   'GRID_BROWSE_ = False
   txtBANK_ID.Enabled = True
End Sub

Private Sub gridMain_RowColChange()
   'GRID_BROWSE_ = True
   populateControl Me, gridMain
End Sub

Private Sub Label9_Click()
    
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
   gridMain.Rows = 2

   populateGridDefineHeader conBank, sSQLQuery_, gridMain, szHeader

   gridMain.RowHeight(0) = 0
   gridMain.ColWidth(0) = "1000"
   gridMain.ColWidth(1) = "1200"
   gridMain.ColWidth(2) = "2000"
   gridMain.ColWidth(3) = "2000"
   gridMain.ColWidth(4) = 0
   gridMain.ColWidth(5) = 0
   gridMain.ColWidth(6) = 0
   gridMain.ColWidth(7) = 0
   gridMain.ColWidth(8) = "1500"
   conBank.Close
   Set conBank = Nothing
End Sub
    

Private Sub lblBankID_Click()
    'issue 0559: Bank Details not sorting correctly
'Written by anol 14 May 2015
    Dim sSQLQuery_ As String, szHeader As String
   Dim conBank As New ADODB.Connection

   conBank.Open getConnectionString
   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
                "FROM tlbBank Order by cint(BANK_ID) "

   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
   gridMain.Rows = 2

   populateGridDefineHeader conBank, sSQLQuery_, gridMain, szHeader

   gridMain.RowHeight(0) = 0
   gridMain.ColWidth(0) = "1000"
   gridMain.ColWidth(1) = "1200"
   gridMain.ColWidth(2) = "2000"
   gridMain.ColWidth(3) = "2000"
   gridMain.ColWidth(4) = 0
   gridMain.ColWidth(5) = 0
   gridMain.ColWidth(6) = 0
   gridMain.ColWidth(7) = 0
   gridMain.ColWidth(8) = "1500"
   conBank.Close
   Set conBank = Nothing

End Sub

Private Sub lblbranch_Click()
      'issue 0559: Bank Details not sorting correctly
'Written by anol 14 May 2015
    Dim sSQLQuery_ As String, szHeader As String
   Dim conBank As New ADODB.Connection

   conBank.Open getConnectionString
   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
                "FROM tlbBank Order by BANK_BRANCH "

   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
   gridMain.Rows = 2

   populateGridDefineHeader conBank, sSQLQuery_, gridMain, szHeader

   gridMain.RowHeight(0) = 0
   gridMain.ColWidth(0) = "1000"
   gridMain.ColWidth(1) = "1200"
   gridMain.ColWidth(2) = "2000"
   gridMain.ColWidth(3) = "2000"
   gridMain.ColWidth(4) = 0
   gridMain.ColWidth(5) = 0
   gridMain.ColWidth(6) = 0
   gridMain.ColWidth(7) = 0
   gridMain.ColWidth(8) = "1500"
   conBank.Close
   Set conBank = Nothing
End Sub

Private Sub lblSortCode_Click()
     'issue 0559: Bank Details not sorting correctly
'Written by anol 14 May 2015
    Dim sSQLQuery_ As String, szHeader As String
   Dim conBank As New ADODB.Connection

   conBank.Open getConnectionString
   sSQLQuery_ = "SELECT BANK_ID, SORT_CODE, BANK_BRANCH, BANK_NAME, BANK_ADDRESS1, " & _
                  "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4, BANK_POST_CODE " & _
                "FROM tlbBank Order by SORT_CODE "

   szHeader$ = "<BANK_ID|<SORT_CODE|<BANK_BRANCH|<BANK_NAME|<BANK_ADDRESS1|<BANK_ADDRESS2|<BANK_ADDRESS3|<BANK_ADDRESS4|<BANK_POST_CODE"
   gridMain.Rows = 2

   populateGridDefineHeader conBank, sSQLQuery_, gridMain, szHeader

   gridMain.RowHeight(0) = 0
   gridMain.ColWidth(0) = "1000"
   gridMain.ColWidth(1) = "1200"
   gridMain.ColWidth(2) = "2000"
   gridMain.ColWidth(3) = "2000"
   gridMain.ColWidth(4) = 0
   gridMain.ColWidth(5) = 0
   gridMain.ColWidth(6) = 0
   gridMain.ColWidth(7) = 0
   gridMain.ColWidth(8) = "1500"
   conBank.Close
   Set conBank = Nothing
End Sub

Private Sub txtBANK_ADDRESS1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBANK_ADDRESS2.SetFocus
    End If
End Sub

Private Sub txtBANK_ADDRESS2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBANK_ADDRESS3.SetFocus
    End If
End Sub

Private Sub txtBANK_ADDRESS3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBANK_ADDRESS4.SetFocus
    End If
End Sub

Private Sub txtBANK_ADDRESS4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBANK_POST_CODE.SetFocus
    End If
End Sub

Private Sub txtBANK_BRANCH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBANK_ADDRESS1.SetFocus
    End If
End Sub

Private Sub txtBANK_NAME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtBANK_BRANCH.SetFocus
    End If
End Sub

Private Sub txtBANK_POST_CODE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
End Sub

Private Sub txtSORT_CODE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBANK_NAME.SetFocus
   End If
   BankSCTextKeyPress txtSORT_CODE, KeyAscii
   
End Sub

