VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLeaseReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lease Details Report"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11580
   Icon            =   "frmLeaseReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   4320
      Top             =   2640
      Width           =   3255
   End
   Begin MSForms.ComboBox cboLeseTo 
      Height          =   315
      Left            =   4320
      TabIndex        =   14
      Top             =   2160
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   3
      ColumnCount     =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;0"
   End
   Begin MSForms.ComboBox cboLeseFrom 
      Height          =   315
      Left            =   960
      TabIndex        =   13
      Top             =   2160
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   3
      ColumnCount     =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;0"
   End
   Begin MSForms.ComboBox cboUnitTo 
      Height          =   315
      Left            =   4320
      TabIndex        =   12
      Top             =   1680
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   3
      ColumnCount     =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;0"
   End
   Begin MSForms.ComboBox cboUnitFrom 
      Height          =   315
      Left            =   960
      TabIndex        =   11
      Top             =   1680
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   3
      ColumnCount     =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;0"
   End
   Begin MSForms.ComboBox cboPropTo 
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   1200
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   3
      ColumnCount     =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;0"
   End
   Begin MSForms.ComboBox cboPropFrom 
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   3
      ColumnCount     =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;0"
   End
   Begin MSForms.ComboBox cboClientTo 
      Height          =   315
      Left            =   4320
      TabIndex        =   8
      Top             =   720
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   2
      ColumnCount     =   2
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0"
   End
   Begin MSForms.ComboBox cboClientFrom 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   720
      Width           =   3255
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   2
      ColumnCount     =   2
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Property "
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "To"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "From"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Generate the Lease details report"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmLeaseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dataClient() As String, dataProp() As String, dataUnit() As String, dataLese() As String

Private Sub cboPropTo_Click()
   If cboPropFrom.text = "" Then Exit Sub

'  Relist Unit and lessee combo
   RelistUL
End Sub

Private Sub cboUnitFrom_Click()
   If cboUnitTo.text = "" Then Exit Sub

'  Relist lessee combo
   RelistL
End Sub

Private Sub cboUnitTo_Click()
   If cboUnitFrom.text = "" Then Exit Sub

'  Relist lessee combo
   RelistL
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdReport_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim szIDList As String, i As Integer
   Dim iFromIdx As Integer, iToIdx As Integer, iTemp As Integer

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LeaseDetailsReport.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   
   If cboLeseFrom.text <> "" And cboLeseTo.text <> "" Then
      Report.ParameterFields(1).AddCurrentValue False
      Report.ParameterFields(2).AddCurrentValue False
      Report.ParameterFields(3).AddCurrentValue False
      Report.ParameterFields(4).AddCurrentValue True

      iFromIdx = cboLeseFrom.ListIndex
      iToIdx = cboLeseTo.ListIndex
      If iFromIdx > iToIdx Then
         iTemp = iFromIdx
         iFromIdx = iToIdx
         iToIdx = iTemp
      End If

      szIDList = ""
      For i = iFromIdx To iToIdx
         szIDList = "'" & cboLeseFrom.Column(0, i) & "', " & szIDList
      Next i
      szIDList = Left(szIDList, Len(szIDList) - 2)

   ElseIf cboUnitFrom.text <> "" And cboUnitTo.text <> "" Then
      Report.ParameterFields(1).AddCurrentValue False
      Report.ParameterFields(2).AddCurrentValue False
      Report.ParameterFields(3).AddCurrentValue True
      Report.ParameterFields(4).AddCurrentValue False

      iFromIdx = cboUnitFrom.ListIndex
      iToIdx = cboUnitTo.ListIndex
      If iFromIdx > iToIdx Then
         iTemp = iFromIdx
         iFromIdx = iToIdx
         iToIdx = iTemp
      End If

      szIDList = ""
      For i = iFromIdx To iToIdx
         szIDList = cboUnitFrom.Column(0, i) & ", " & szIDList
      Next i
      szIDList = Left(szIDList, Len(szIDList) - 2)
   
   ElseIf cboPropFrom.text <> "" And cboPropTo.text <> "" Then
      Report.ParameterFields(1).AddCurrentValue False
      Report.ParameterFields(2).AddCurrentValue True
      Report.ParameterFields(3).AddCurrentValue False
      Report.ParameterFields(4).AddCurrentValue False

      iFromIdx = cboPropFrom.ListIndex
      iToIdx = cboPropTo.ListIndex
      If iFromIdx > iToIdx Then
         iTemp = iFromIdx
         iFromIdx = iToIdx
         iToIdx = iTemp
      End If

      szIDList = ""
      For i = iFromIdx To iToIdx
'         szIDList = "'" & dataProp(0, i) & "', " & szIDList
         szIDList = "'" & cboPropFrom.Column(0, i) & "', " & szIDList
      Next i
      szIDList = Left(szIDList, Len(szIDList) - 2)
   Else
      Report.ParameterFields(1).AddCurrentValue True
      Report.ParameterFields(2).AddCurrentValue False
      Report.ParameterFields(3).AddCurrentValue False
      Report.ParameterFields(4).AddCurrentValue False

      szIDList = ""
      If cboClientFrom.text <> "" And cboClientTo.text <> "" Then
         iFromIdx = cboClientFrom.ListIndex
         iToIdx = cboClientTo.ListIndex
         If iFromIdx > iToIdx Then
            iTemp = iFromIdx
            iFromIdx = iToIdx
            iToIdx = iTemp
         End If

         For i = iFromIdx To iToIdx
'            szIDList = "'" & dataClient(0, i) & "', " & szIDList
            szIDList = "'" & cboClientFrom.Column(0, i) & "', " & szIDList
         Next i
      Else
         For i = 0 To cboClientTo.ListCount - 1
'            szIDList = "'" & dataClient(0, i) & "', " & szIDList
            szIDList = "'" & cboClientFrom.Column(0, i) & "', " & szIDList
         Next i
      End If
      
      szIDList = Left(szIDList, Len(szIDList) - 2)
   End If
   Report.ParameterFields(5).AddCurrentValue szIDList

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub Form_Load()
   Me.Width = 7785
   Me.Height = 3855
   Me.BackColor = MODULEBACKCOLOR
   
   Dim i As Integer, j As Integer
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   
   LoadAllData adoConn

   adoConn.Close
   Set adoConn = Nothing
   
   LoadLese dataLese
   LoadUnit dataUnit
   LoadProperty dataProp
   LoadClient
End Sub

Private Sub LoadLese(ByRef dataData() As String)
   Dim i As Integer, X As Integer, Y As Integer

   X = UBound(dataData, 1)
   Y = UBound(dataData, 2)
   
   For i = 0 To Y
      If dataData(0, i) = "" Then
         Exit For
      End If
   Next i

   ReDim Preserve dataData(X, i) As String

   cboLeseFrom.Clear
   cboLeseTo.Clear
   
   cboLeseFrom.Column() = dataData()
   cboLeseTo.Column() = dataData()
End Sub

Private Sub LoadUnit(ByRef dataData() As String)
   cboUnitFrom.Clear
   cboUnitTo.Clear

   cboUnitFrom.Column() = dataData()
   cboUnitTo.Column() = dataData()
End Sub

Private Sub LoadProperty(ByRef dataData() As String)
   cboPropFrom.Clear
   cboPropTo.Clear

   cboPropFrom.Column() = dataData()
   cboPropTo.Column() = dataData()
End Sub

Private Sub LoadClient()
   cboClientFrom.Clear
   cboClientTo.Clear

   cboClientFrom.Column() = dataClient()
   cboClientTo.Column() = dataClient()
End Sub

Private Sub cboPropFrom_Click()
   If cboPropTo.text = "" Then Exit Sub

'  Relist Unit and lessee combo
   RelistUL
End Sub

Private Sub cboClientTo_Click()
   If cboClientFrom.text = "" Then Exit Sub

'  Relist Property, Unit and lessee combo
   RelistPUL
End Sub

Private Sub cboClientFrom_Click()
   If cboClientTo.text = "" Then Exit Sub

'  Relist Property, Unit and lessee combo
   RelistPUL
End Sub

Private Sub RelistL()
   Dim i As Integer, j As Integer, r As Integer, c As Integer, K As Integer
   Dim dataData() As String, iFromIdx As Integer, iToIdx As Integer, iTemp As Integer

'  Lessee
   r = UBound(dataLese, 1)
   c = UBound(dataLese, 2)

   ReDim dataData(r, c) As String

   iFromIdx = cboUnitFrom.ListIndex
   iToIdx = cboUnitTo.ListIndex
   If iFromIdx > iToIdx Then
      iTemp = iFromIdx
      iFromIdx = iToIdx
      iToIdx = iTemp
   End If
   
   K = 0
   For i = 0 To c
      If InTheList(dataLese(1, i), cboUnitFrom, iFromIdx, iToIdx) Then
         For j = 0 To r
            dataData(j, K) = dataLese(j, i)
         Next j
         K = K + 1
      End If
   Next i
   LoadLese dataData()

   cboLeseFrom.ListIndex = -1
   cboLeseTo.ListIndex = -1
End Sub

Private Sub RelistUL()
   Dim i As Integer, j As Integer, r As Integer, c As Integer, K As Integer
   Dim dataData() As String, iFromIdx As Integer, iToIdx As Integer, iTemp As Integer

'  Unit
   r = UBound(dataUnit, 1)
   c = UBound(dataUnit, 2)

   ReDim dataData(r, c) As String

   iFromIdx = cboPropFrom.ListIndex
   iToIdx = cboPropTo.ListIndex
   If iFromIdx > iToIdx Then
      iTemp = iFromIdx
      iFromIdx = iToIdx
      iToIdx = iTemp
   End If
   
   K = 0
   For i = 0 To c
      If InTheList(dataUnit(1, i), cboPropFrom, iFromIdx, iToIdx) Then
         For j = 0 To r
            dataData(j, K) = dataUnit(j, i)
         Next j
         K = K + 1
      End If
   Next i
   LoadUnit dataData()
   
'  Lessee
   r = UBound(dataLese, 1)
   c = UBound(dataLese, 2)

   ReDim dataData(r, c) As String

   K = 0
   For i = 0 To c
      If InTheList(dataLese(1, i), cboUnitFrom, 0, cboUnitTo.ListCount - 1) Then
         For j = 0 To r
            dataData(j, K) = dataLese(j, i)
         Next j
         K = K + 1
      End If
   Next i
   LoadLese dataData()
   
   cboUnitFrom.ListIndex = -1
   cboUnitTo.ListIndex = -1
   cboLeseFrom.ListIndex = -1
   cboLeseTo.ListIndex = -1
End Sub

Private Sub RelistPUL()
   Dim i As Integer, j As Integer, r As Integer, c As Integer, K As Integer
   Dim dataData() As String, iFromIdx As Integer, iToIdx As Integer, iTemp As Integer

'  Property
   r = UBound(dataProp, 1)
   c = UBound(dataProp, 2)

   ReDim dataData(r, c) As String

   iFromIdx = cboClientFrom.ListIndex
   iToIdx = cboClientTo.ListIndex
   If iFromIdx > iToIdx Then
      iTemp = iFromIdx
      iFromIdx = iToIdx
      iToIdx = iTemp
   End If

   K = 0
   For i = 0 To c
      If InTheList(dataProp(1, i), cboClientFrom, iFromIdx, iToIdx) Then
         For j = 0 To r
            dataData(j, K) = dataProp(j, i)
         Next j
         K = K + 1
      End If
   Next i
   LoadProperty dataData()

'  Unit
   r = UBound(dataUnit, 1)
   c = UBound(dataUnit, 2)

   ReDim dataData(r, c) As String

   iFromIdx = 0
   iToIdx = cboPropTo.ListCount - 1
   If iFromIdx > iToIdx Then
      iTemp = iFromIdx
      iFromIdx = iToIdx
      iToIdx = iTemp
   End If

   K = 0
   For i = 0 To c
      If InTheList(dataUnit(1, i), cboPropFrom, iFromIdx, iToIdx) Then
         For j = 0 To r
            dataData(j, K) = dataUnit(j, i)
         Next j
         K = K + 1
      End If
   Next i
   LoadUnit dataData()
   
'  Lessee
   r = UBound(dataLese, 1)
   c = UBound(dataLese, 2)

   ReDim dataData(r, c) As String

   iFromIdx = 0
   iToIdx = cboUnitFrom.ListCount - 1
   If iFromIdx > iToIdx Then
      iTemp = iFromIdx
      iFromIdx = iToIdx
      iToIdx = iTemp
   End If

   K = 0
   For i = 0 To c
      If InTheList(dataLese(1, i), cboUnitFrom, iFromIdx, iToIdx) Then
         For j = 0 To r
            dataData(j, K) = dataLese(j, i)
         Next j
         K = K + 1
      End If
   Next i
   LoadLese dataData()

   cboPropFrom.ListIndex = -1
   cboPropTo.ListIndex = -1
   cboUnitFrom.ListIndex = -1
   cboUnitTo.ListIndex = -1
   cboLeseFrom.ListIndex = -1
   cboLeseTo.ListIndex = -1
End Sub

Private Function InTheList(szLooking As String, ctLookingInto As MSForms.ComboBox, iStPos As Integer, iEnPos As Integer) As Boolean
   Dim i As Integer

   InTheList = False
   For i = iStPos To iEnPos
      If ctLookingInto.ListIndex = -1 Then ctLookingInto.ListIndex = i
      If szLooking = ctLookingInto.Column(0, i) Then
         InTheList = True
         Exit For
      End If
   Next i
End Function

Private Sub LoadAllData(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim szSQL As String

   On Error GoTo ErrorHandler

'  Load Clients' data in the memory

   szSQL = "SELECT DISTINCT ClientID, ClientName " & _
           "FROM Client " & _
           "ORDER BY ClientName;"
           
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count
   ReDim dataClient(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           dataClient(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   adoRst.Close
   
'  Load Properties' data in the memory
   
   szSQL = "SELECT DISTINCT PropertyID, ClientID, PropertyName " & _
           "FROM Property;" '& _
           "ORDER BY PropertyName;"
           
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count
   ReDim dataProp(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           dataProp(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   adoRst.Close
   
'  Load Units' data in the memory
   
   szSQL = "SELECT DISTINCT UnitNumber, PropertyID, UnitName " & _
           "FROM Units " '& _
           "ORDER BY UnitName;"
           
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count
   ReDim dataUnit(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           dataUnit(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   adoRst.Close
   
'  Load Lessees' data in the memory
   
   szSQL = "SELECT DISTINCT Tenants.SageAccountNumber, UnitNumber, Name " & _
           "FROM Tenants, LeaseDetails " & _
           "WHERE Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber;" '& _
           "ORDER BY Name;"
           
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count
   ReDim dataLese(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           dataLese(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

