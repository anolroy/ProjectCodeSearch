VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Property"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "frmProperty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtProPostCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      TabIndex        =   17
      Top             =   2340
      Width           =   1455
   End
   Begin VB.TextBox txtProAddressLine3 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      TabIndex        =   16
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtProAddressLine2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      TabIndex        =   15
      Top             =   1500
      Width           =   2775
   End
   Begin VB.TextBox txtProAddressLine1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      TabIndex        =   13
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtPropertyName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      TabIndex        =   6
      Top             =   660
      Width           =   2775
   End
   Begin VB.TextBox txtPropertyID 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6EDFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   1770
      TabIndex        =   4
      Top             =   7380
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6150
      TabIndex        =   3
      Top             =   7380
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5055
      TabIndex        =   2
      Top             =   7380
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   7380
      Width           =   1035
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2865
      TabIndex        =   0
      Top             =   7380
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   375
      Left            =   60
      Top             =   4140
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
   Begin MSAdodcLib.Adodc adoClient 
      Height          =   375
      Left            =   60
      Top             =   3840
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
      Caption         =   "Client"
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
   Begin MSDataListLib.DataCombo cboClientID 
      Height          =   315
      Left            =   1380
      TabIndex        =   7
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMain 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   3180
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox cboTest 
      Height          =   315
      Left            =   4560
      TabIndex        =   19
      Top             =   1020
      Width           =   3450
      VariousPropertyBits=   1820346395
      DisplayStyle    =   3
      Size            =   "6085;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   3
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
      Caption         =   "Post Code:"
      Height          =   255
      Left            =   300
      TabIndex        =   18
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Address:"
      Height          =   255
      Left            =   300
      TabIndex        =   14
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Client:"
      Height          =   255
      Left            =   300
      TabIndex        =   12
      Top             =   300
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   300
      TabIndex        =   11
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Property ID:"
      Height          =   255
      Left            =   4980
      TabIndex        =   10
      Top             =   300
      Width           =   1035
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H009F0258&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F9C6CD&
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   8715
   End
End
Attribute VB_Name = "frmProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROPERTY_NEW_ENTRY_ As Boolean
Dim GRID_BROWSE_ As Boolean
''

Private Sub cboClientID_Change()
'If Not GRID_BROWSE_ Then
    LoadGrid
'End If
End Sub


Private Sub cboTest_Click()
'MsgBox cboTest.Column(0) & " "
End Sub

Private Sub cmdCancel_Click()
ComponentEnableMode frmProperty, NewEntryMode, gridMain
ComponentEnableMode frmProperty, DefaultMode, gridMain
cboClientID.Enabled = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If gridMain.TextMatrix(gridMain.Row, 0) = "" Then
    MsgBox "Please select a record to continue.", vbInformation, "Edit record"
    Exit Sub
End If
PROPERTY_NEW_ENTRY_ = False
ComponentEnableMode frmProperty, EditMode, gridMain
cboClientID.Enabled = False
txtPropertyName.SetFocus
End Sub

Private Sub cmdNew_Click()
If cboClientID.text = "" Then
    MsgBox "Please select a client to continue.", vbInformation, "New Property"
    Exit Sub
End If

PROPERTY_NEW_ENTRY_ = True
ComponentEnableMode frmProperty, NewEntryMode, gridMain
cboClientID.Enabled = False
txtPropertyName.SetFocus
End Sub

Private Sub cmdSave_Click()
Dim sSQLQuery_ As String, sFilter As String

adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

If PROPERTY_NEW_ENTRY_ = True Then
    sFilter = " WHERE PROPERTYID = '" & txtPropertyID.text & "' "
Else
    sFilter = " WHERE PROPERTYID = '" & txtPropertyID.text & "' "
End If

sSQLQuery_ = "SELECT * " & _
              "FROM PROPERTY" & sFilter

adoMain.RecordSource = sSQLQuery_
adoMain.CommandType = adCmdText
adoMain.Refresh

If PROPERTY_NEW_ENTRY_ = True Then
    'adoMain.Recordset.AddNew
    Add_NoQuery frmProperty, adoMain
Else
    Update_NoQuery frmProperty, adoMain
End If

ComponentEnableMode frmProperty, DefaultMode, gridMain
cboClientID.Enabled = True
LoadGrid
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub Form_Load()

Me.Top = 50
Me.Left = 50

PopulateClient
ComponentEnableMode frmProperty, DefaultMode, gridMain
adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
LoadGrid

cboClientID.Enabled = True

gridMain.ColWidth(0) = "10"
gridMain.ColWidth(1) = "10"
gridMain.ColWidth(2) = "1700"
gridMain.ColWidth(3) = "1700"
gridMain.ColWidth(4) = "1700"
gridMain.ColWidth(5) = "1700"
gridMain.ColWidth(6) = "1700"
gridMain.ColWidth(7) = "1400"
End Sub

Public Function PopulateClient()
Dim sSQLQuery_ As String

adoClient.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
'adoClient.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & Adsn

sSQLQuery_ = "SELECT TenantId, Name, BillPostCode " & _
              "FROM Tenants "

adoClient.RecordSource = sSQLQuery_
adoClient.CommandType = adCmdText
adoClient.Refresh

Dim TotalRow, TotalCol As Integer

TotalRow = adoClient.Recordset.RecordCount
TotalCol = adoClient.Recordset.Fields.Count

Dim data() As String

ReDim data(TotalCol, TotalRow) As String

Dim i, j As Integer

For i = 0 To adoClient.Recordset.RecordCount - 1
    For j = 0 To adoClient.Recordset.Fields.Count - 1
    'cboTest.AddItem adoClient.Recordset.Fields(1)
        data(j, i) = adoClient.Recordset.Fields(j)
    Next j
    adoClient.Recordset.MoveNext
Next i
'
cboTest.Column() = data()
'populateCombo sSQLQuery_, cboTest

End Function

Public Function LoadGrid()
Dim sSQLQuery_ As String

sSQLQuery_ = "SELECT CLIENT.CLIENTNAME, PROPERTY.PROPERTYID, PROPERTY.PROPERTYNAME, " & _
              "PROPERTY.PROADDRESSLINE1, PROPERTY.PROADDRESSLINE2,PROPERTY.PROADDRESSLINE3, " & _
              "PROPOSTCODE " & _
              "FROM CLIENT, PROPERTY " & _
              "WHERE PROPERTY.CLIENTID = '" & cboClientID.BoundText & "' " & _
              "AND CLIENT.CLIENTID = PROPERTY.CLIENTID "
              
adoMain.RecordSource = sSQLQuery_
adoMain.CommandType = adCmdText
adoMain.Refresh

populateGrid frmProperty, adoMain, gridMain
End Function

Private Sub Form_Unload(Cancel As Integer)
frmMMain.fraCmdButton.Enabled = True
Unload Me
End Sub



Private Sub gridMain_LostFocus()
GRID_BROWSE_ = False
'ComponentEnableMode frmProperty, GridLostFocus, gridMain
cboClientID.Enabled = True
End Sub

Private Sub gridMain_RowColChange()

GRID_BROWSE_ = True
ComponentEnableMode frmProperty, GridRowOnSelection, gridMain
populateControl frmProperty, gridMain
End Sub

Public Function GeneratePropertyID() As String

Dim conPropertyID As New RDO.rdoConnection
Dim rstPropertyID As rdoResultset
Dim sSQLQuery_ As String
Dim MAX_Property_ As String
Dim PROPERTY_ID_ As String

'On Error Resume Next
'Set the RDO Connections to the dataset
conPropertyID.Connect = "DSN=" & Adsn & ";UID=;PWD="
conPropertyID.CursorDriver = rdUseIfNeeded
conPropertyID.EstablishConnection rdDriverNoPrompt

'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
sSQLQuery_ = "SELECT MAX(RIGHT(PROPERTY.PROPERTYID,2)) + 1 AS  MAX_PROPERTYID " & _
    " " & _
    "From Property " & _
    "WHERE LEFT(PROPERTY.PropertyID,2) = LEFT(TRIM('" & txtPropertyName.text & "'),2)"
'Debug.Print sSQLQuery_

Set rstPropertyID = conPropertyID.OpenResultset(sSQLQuery_, rdOpenStatic, rdConcurReadOnly)
'MsgBox sUnitNumber

If rstPropertyID.EOF Or rstPropertyID.BOF Then
    MAX_Property_ = "1"
End If

While Not rstPropertyID.EOF

    MAX_Property_ = IIf(IsNull(rstPropertyID!MAX_PROPERTYID), "1", rstPropertyID!MAX_PROPERTYID)
    rstPropertyID.MoveNext

Wend

GeneratePropertyID = UCase(Left(txtPropertyName.text, 2)) & Lpad(MAX_Property_, "0", 2)


rstPropertyID.Close
conPropertyID.Close
Set rstPropertyID = Nothing
Set conPropertyID = Nothing

End Function


Private Sub txtPropertyName_LostFocus()
    If PROPERTY_NEW_ENTRY_ And txtPropertyID.text <> "" Then
        txtPropertyID.text = GeneratePropertyID
    End If
End Sub
