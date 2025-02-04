VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDCTypesPre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demand Types/Charge Types/Rent Payable"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCTypesPre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   930
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3675
      TabIndex        =   3
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   765
   End
   Begin MSForms.ComboBox cboClientList 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3645
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "6429;556"
      BoundColumn     =   0
      TextColumn      =   1
      ColumnCount     =   8
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboPropertyList 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   3645
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "6429;556"
      BoundColumn     =   0
      TextColumn      =   1
      ColumnCount     =   3
      ListRows        =   20
      cColumnInfo     =   4
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1411;4233;-1;1411"
   End
End
Attribute VB_Name = "frmDCTypesPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public szMenu As String
Private DataProperties() As String

Private Sub cboPropertyList_Change()
   On Error GoTo XXX

   If cboPropertyList.Column(0) <> "ALL" Then
      cboClientList.ListIndex = ComboListIndex(cboClientList, cboPropertyList.Column(4), 0)
   End If
XXX:
End Sub

Private Function ComboListIndex(conCombo As Control, szText As String, iCol As Integer) As Integer
   Dim i As Integer
   
   For i = 0 To conCombo.ListCount - 1
      
      If conCombo.Column(iCol, i) = szText Then
         ComboListIndex = i
         Exit Function
      End If
   Next i
End Function

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Me.Hide

   If szMenu = "DEMAND_TYPE" Then
'      frmDCTypes.fraDemandType.Visible = True
'      frmDCTypes.fraChargeType.Visible = False
'      frmDCTypes.fraPayableType.Visible = False
'      frmDCTypes.Caption = "Demand Type"
'      frmDCTypes.lblClient.Caption = cboClientList.Column(1)
'      frmDCTypes.lblProperty.Caption = cboPropertyList.Column(1)
      Load frmDemandTypes
      frmDemandTypes.Show
   End If
   If szMenu = "CHARGE_TYPE" Then
      Load frmDCTypes
      frmDCTypes.fraDemandType.Visible = False
      frmDCTypes.fraChargeType.Visible = True
      frmDCTypes.fraPayableType.Visible = False
      frmDCTypes.Caption = "Charge Type"
      frmDCTypes.Show
   End If
   If szMenu = "PAYABLE_TYPE" Then
      Load frmDCTypes
      frmDCTypes.fraDemandType.Visible = False
      frmDCTypes.fraChargeType.Visible = False
      frmDCTypes.fraPayableType.Visible = True
      frmDCTypes.Caption = "Payable Type"
      frmDCTypes.Show
   End If

End Sub

Private Sub Form_Activate()
   cboClientList.ListIndex = 0
   cboPropertyList.ListIndex = 0
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection

   Me.Height = 2115
   Me.Width = 4860
   Me.Top = 500
   Me.Left = 500
   Me.BackColor = MODULEBACKCOLOR

   adoConn.Open getConnectionString

   PrepareList adoConn, cboClientList, cboPropertyList

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cboClientList_CLICK()
   Dim szSQL   As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   On Error Resume Next

   For i = 0 To UBound(DataProperties, 2)
      If cboClientList.Column(0) = "ALL" Then
         TotalRow = TotalRow + 1
      Else
         If DataProperties(4, i) = cboClientList.Column(0) Then
            TotalRow = TotalRow + 1
         End If
      End If
   Next i

   If TotalRow = 0 Then Exit Sub

   ReDim Data(UBound(DataProperties, 1), TotalRow)

   If cboClientList.Column(0) = "ALL" Then
      For i = 0 To UBound(DataProperties, 2)
         For j = 0 To UBound(DataProperties, 1) - 1
            Data(j, i) = DataProperties(j, i)
         Next j
      Next i
   Else
      TotalRow = 0
      For i = 0 To UBound(DataProperties, 2)
         If DataProperties(4, i) = cboClientList.Column(0) Then
            For j = 0 To UBound(DataProperties, 1)
               Data(j, TotalRow) = DataProperties(j, i)
            Next j
            TotalRow = TotalRow + 1
         End If
      Next i
   End If

   cboPropertyList.Clear
   cboPropertyList.Column() = Data()
   cboPropertyList.ListIndex = 0
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

   cboClient.Column() = Data()
   adoRst.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode, ClientID " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String
   ReDim DataProperties(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   DataProperties(0, 0) = "ALL"
   DataProperties(1, 0) = "All Properties"
   For i = 1 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
         DataProperties(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i

   cboProperty.Column() = Data()
NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number
End Sub
