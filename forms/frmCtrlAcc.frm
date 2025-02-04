VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCtrlAcc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Control Account"
   ClientHeight    =   2115
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCtrlAcc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCtrlName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   405
   End
   Begin MSForms.ComboBox cboType 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4895;503"
      TextColumn      =   2
      ColumnCount     =   3
      ListRows        =   20
      cColumnInfo     =   3
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527;0"
   End
   Begin VB.Label lblClient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   2745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control Account:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label lblClient1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmCtrlAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AddNew As Boolean

Private Sub CadataelButton_Click()
   Unload Me
   frmOptions.Enabled = True
End Sub

Private Sub CancelButton_Click()
   Unload Me
   frmOptions.ControlHanlding DefaultMode
End Sub

Private Sub cboType_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then CadataelButton_Click
End Sub

Private Sub Form_Load()
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim TotalRow As Long, TotalCol As Long
   Dim data() As String

   adoConn.Open getConnectionString

   szSQL = "SELECT Code, Value " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = 'CAT' AND CODE IN ('P', 'S');"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.count
   ReDim data(TotalCol - 1, TotalRow - 1) As String

   Dim i As Integer, j As Integer

   For i = 0 To adoRST.RecordCount - 1
      For j = 0 To adoRST.Fields.count - 1
         data(j, i) = IIf(IsNull(adoRST.Fields(j)), "", adoRST.Fields(j))
      Next j
      adoRST.MoveNext
   Next i

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   cboType.Column() = data()
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmOptions.Enabled = True
End Sub

Private Sub OKButton_Click()
   If txtCtrlName.text = "" Then
      ShowMsgInTaskBar "Enter control account name", "Y", "N"
      txtCtrlName.SetFocus
      Exit Sub
   End If
   If IsNull(cboType.Value) Then
      ShowMsgInTaskBar "Please select the type", "Y", "N"
      cboType.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRST  As New ADODB.Recordset
   Dim X       As Integer
   Dim szID    As String
   Dim szSQL   As String

   adoConn.Open getConnectionString

   If AddNew Then
      adoRST.Open "SELECT CAName " & _
                  "FROM NominalLedger " & _
                  "WHERE ClientID = '" & frmOptions.txtClientList.text & "' AND " & _
                        "CAName = '" & txtCtrlName.text & "' AND " & _
                        "(CAType = 'S' OR CAType = 'P');", adoConn, adOpenStatic, adLockReadOnly
      If Not adoRST.EOF Then
         ShowMsgInTaskBar "Already the code account name has been used", "Y", "N"
         adoRST.Close
         Set adoRST = Nothing
         adoConn.Close
         Set adoConn = Nothing
         Exit Sub
      End If
      adoRST.Close
   End If

   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   With frmOptions
      If AddNew Then
         .flxTransactionTypes.AddItem ""
         X = .flxTransactionTypes.Rows - 1
         .flxTransactionTypes.TextMatrix(X, 0) = szID
         .flxTransactionTypes.TextMatrix(X, 5) = .txtClientList.text
         .flxTransactionTypes.TextMatrix(X, 6) = False
         .flxTransactionTypes.TextMatrix(X, 7) = "NO"
      Else
         X = .flxTransactionTypes.row
      End If
      .flxTransactionTypes.TextMatrix(X, 1) = txtCtrlName.text
      .flxTransactionTypes.TextMatrix(X, 2) = cboType.text
      .flxTransactionTypes.TextMatrix(X, 8) = cboType.Value
      .flxTransactionTypes.TextMatrix(X, 9) = "N"

      Unload Me
   End With
End Sub

Private Sub OKButton_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then CadataelButton_Click
End Sub

Private Sub txtCtrlName_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then CadataelButton_Click
End Sub
