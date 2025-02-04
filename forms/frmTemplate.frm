VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letter Template"
   ClientHeight    =   13500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTemplate.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13500
   ScaleWidth      =   13950
   Begin VB.Frame fraLetterBody 
      BorderStyle     =   0  'None
      Caption         =   "Subject"
      Height          =   4215
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox txtOurRef 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1455
         MaxLength       =   200
         TabIndex        =   48
         Top             =   630
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8685
         TabIndex        =   7
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   1440
         TabIndex        =   14
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   10
         Top             =   600
         Width           =   5175
      End
      Begin VB.CommandButton cmdSelectdiff 
         Caption         =   "&Cancel"
         Height          =   510
         Index           =   0
         Left            =   5880
         TabIndex        =   15
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox txtTemplateName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtSalutation 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         TabIndex        =   9
         Top             =   2355
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtSPDate 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPosition 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   13
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtSender 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtBody 
         DataSource      =   "Adodc1"
         Height          =   2175
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   960
         Width           =   9855
      End
      Begin MSAdodcLib.Adodc adcTemplateSave1 
         Height          =   375
         Left            =   0
         Top             =   3720
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Our Ref:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   49
         Top             =   630
         Width           =   555
      End
      Begin MSForms.CommandButton cmdpropertyunit 
         Height          =   510
         Left            =   9360
         TabIndex        =   45
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
         Caption         =   "Select Units >>"
         Size            =   "3413;900"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Template Type:"
         Height          =   195
         Left            =   5400
         TabIndex        =   44
         Top             =   240
         Width           =   1065
      End
      Begin MSForms.ComboBox cmbTempType 
         Height          =   315
         Left            =   6480
         TabIndex        =   6
         Top             =   195
         Width           =   2205
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "3889;556"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   30
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salutation:"
         Height          =   195
         Index           =   0
         Left            =   -195
         TabIndex        =   25
         Top             =   2355
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of letter:"
         Height          =   195
         Left            =   9240
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPosition 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sender's Position:"
         Height          =   195
         Left            =   6480
         TabIndex        =   19
         Top             =   3240
         Width           =   1260
      End
      Begin VB.Label lblSender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sender:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label lblBody 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Body"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Template Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame fraListTemplates 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   390
         Left            =   9420
         TabIndex        =   47
         Top             =   3760
         Width           =   1935
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Template"
         Height          =   390
         Left            =   3140
         TabIndex        =   1
         Top             =   3760
         Width           =   1935
      End
      Begin VB.CheckBox chkAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   9960
         TabIndex        =   42
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkLogo 
         Caption         =   "Logo,"
         Height          =   255
         Left            =   9240
         TabIndex        =   41
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Add &New Template"
         Height          =   390
         Left            =   6280
         TabIndex        =   2
         Top             =   3760
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Edit Template"
         Height          =   390
         Left            =   120
         TabIndex        =   0
         Top             =   3760
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTemplates 
         Height          =   3285
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   5794
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   15329508
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
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
         _Band(0).Cols   =   10
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Print company "
         Height          =   195
         Index           =   1
         Left            =   8160
         TabIndex        =   43
         Top             =   3720
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   9765
         TabIndex        =   23
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modified by"
         Height          =   195
         Left            =   4380
         TabIndex        =   22
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   195
         Left            =   5730
         TabIndex        =   21
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Template Name"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame fraPropertyUnit 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   720
      TabIndex        =   26
      Top             =   8880
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdLastSelection 
         Caption         =   "Last Selection"
         Height          =   495
         Left            =   4920
         TabIndex        =   40
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddUnitAll 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   39
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdRemUnitAll 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   38
         Top             =   2475
         Width           =   375
      End
      Begin VB.CommandButton cmdRemUnit 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   37
         Top             =   3075
         Width           =   375
      End
      Begin VB.CommandButton cmdAddUnit 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   36
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "Print Letters >>"
         Height          =   510
         Left            =   9480
         TabIndex        =   35
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ListBox lstProperty 
         Height          =   3180
         ItemData        =   "frmTemplate.frx":0442
         Left            =   120
         List            =   "frmTemplate.frx":0444
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   360
         Width           =   3375
      End
      Begin VB.ListBox lstUnit 
         Height          =   3180
         ItemData        =   "frmTemplate.frx":0446
         Left            =   3600
         List            =   "frmTemplate.frx":0448
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   3375
      End
      Begin VB.ListBox lstDisplay 
         Height          =   3180
         ItemData        =   "frmTemplate.frx":044A
         Left            =   7560
         List            =   "frmTemplate.frx":044C
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton cmdback 
         Caption         =   "<< Back"
         Height          =   510
         Left            =   120
         TabIndex        =   29
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Units/Lessees"
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   31
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Units of "
         Height          =   195
         Left            =   3600
         TabIndex        =   28
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bChoice   As Boolean
Private bSaved    As Boolean
Private bValidate As Boolean
Public szLetter   As String

Private Sub cmdAddUnit_Click()
   Dim j As Integer

   If lstUnit.SelCount < 1 Then Exit Sub

   Do
      If lstUnit.Selected(j) Then
         lstDisplay.AddItem lstUnit.text
         lstUnit.RemoveItem lstUnit.ListIndex
         j = j - 1
      End If
      j = j + 1
   Loop While j <= lstUnit.ListCount - 1
End Sub

Private Sub cmdAddUnitAll_Click()
   Dim i As Integer, j As Integer

   If lstUnit.ListCount = 0 Then Exit Sub

   For j = 0 To lstUnit.ListCount - 1
      For i = 0 To lstDisplay.ListCount - 1
         If lstUnit.List(j) = lstDisplay.List(i) Then Exit For
      Next i
      If i = lstDisplay.ListCount Then lstDisplay.AddItem lstUnit.List(j)
   Next j
End Sub

Private Sub cmdBack_Click()
   fraLetterBody.Visible = True
   fraPropertyUnit.Visible = False
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   If flxTemplates.TextMatrix(flxTemplates.row, 0) = "" Then Exit Sub

   If szLetter = "RT" And flxTemplates.TextMatrix(flxTemplates.row, 0) < 4 Then
      ShowMsgInTaskBar "You cannot delete pre-defined reminder template.", "Y", "N"
      Exit Sub
   End If

   If MsgBox("Do you wish to delete the template?", vbQuestion + vbYesNo, "Delete Template") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "DELETE * " & _
           "FROM   Template " & _
           "WHERE  TemplateID = " & flxTemplates.TextMatrix(flxTemplates.row, 0) & ";"
   adoConn.Execute szSQL
   ConfigFlxTemplates
   LoadFlxTemplate adoConn
   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Template has been deleted.", "Y", "P"
'   If flxTemplates.row > 1 Then
'      flxTemplates.RemoveItem flxTemplates.row
'   Else
'      flxTemplates.Clear
'   End If
End Sub

Private Sub cmdLastSelection_Click()
   Dim xSQL As String
    Dim adoConn As New ADODB.Connection
    Dim rsUnit As New ADODB.Recordset
    adoConn.Open getConnectionString
   If (Not bChoice) Then
      xSQL = "Select UnitNumber FROM TemplateUnitSelection " & _
             "Where TemplateID = " & flxTemplates.TextMatrix(flxTemplates.row, 0) & ";"
   
      rsUnit.Open xSQL, adoConn, adOpenKeyset
     
   
      If rsUnit.RecordCount = 0 Then
         ShowMsgInTaskBar "There is no previous selection found.", , "N"
      Else
         lstDisplay.Clear
         While Not rsUnit.EOF
            lstDisplay.AddItem rsUnit.Fields.Item("UnitNumber").Value
            rsUnit.MoveNext
         Wend
      End If
   Else
      ShowMsgInTaskBar "There is no previous selection found.", , "N"
   End If
   adoConn.Close
End Sub

Private Sub cmdNew_Click()
   txtTemplateName.text = ""
   txtSPDate.text = Format(Date, "dd/mm/yyyy")
   txtSalutation.text = "Dear Sir/Madam,"
   txtSubject.text = ""
   txtBody.text = ""
   txtSender.text = ""
   txtPosition.text = ""
   txtOurRef.text = ""
   cmbTempType.Value = szLetter

   fraListTemplates.Visible = False
   fraLetterBody.Visible = True

   bChoice = True
End Sub

Private Sub LoadTempType(szValue As String, adoConn As ADODB.Connection)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRST As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM SecondaryCode " & _
             "WHERE SecondaryCode.PrimaryCode = '" & szValue & "' " & _
             "ORDER BY SecondaryCode.Value;"

   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   cmbTempType.Clear
   i = 0
   While Not adoRST.EOF
      szaData(0, i) = adoRST!c
      szaData(1, i) = adoRST!V
      adoRST.MoveNext
      i = i + 1
   Wend
   adoRST.Close
   Set adoRST = Nothing

   cmbTempType.Column() = szaData()
End Sub

Private Sub cmdNext_Click()
   Dim j As Integer
   If lstDisplay.ListCount < 1 Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, i As Integer, szaUnit() As String
   
   Dim rstRst As New ADODB.Recordset
   Dim conUnit_ As New ADODB.Connection
   Dim rstLessee As New ADODB.Recordset, rstReport As New ADODB.Recordset
   Dim sSQLQuery_ As String
   
   adoConn.Open getConnectionString
   
   If (Not bChoice) Then
      szSQL = "DELETE * FROM TemplateUnitSelection " & _
                   "WHERE TemplateID = " & flxTemplates.TextMatrix(flxTemplates.row, 0) & ";"
      adoConn.Execute szSQL
      
      For j = 0 To lstDisplay.ListCount - 1
         szSQL = "INSERT INTO TemplateUnitSelection (TemplateID, UnitNumber) " & _
                 "VALUES (" & flxTemplates.TextMatrix(flxTemplates.row, 0) & ", '" & lstDisplay.List(j) & "');"
         adoConn.Execute szSQL
      Next j
   End If

   szSQL = "UPDATE Tenants SET TemplatePrint = '';"
   adoConn.Execute szSQL

   conUnit_.Open getConnectionString

   For i = 0 To lstDisplay.ListCount - 1
      szaUnit = Split(lstDisplay.List(i), " \ ")
      szSQL = "UPDATE Tenants, LeaseDetails SET Tenants.TemplatePrint = 'Y' " & _
              "WHERE Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber And " & _
                  "LeaseDetails.UnitNumber = '" & szaUnit(0) & "';"
      adoConn.Execute szSQL
      
      sSQLQuery_ = "SELECT t.SageAccountNumber, t.Name, t.HOAddressLine1, t.HOAddressLine2, t.HOAddressLine3, t.HOAddressLine4, HOPostCode " & _
                "FROM Tenants t, LeaseDetails " & _
                "WHERE t.SageAccountNumber = LeaseDetails.SageAccountNumber And " & _
                  "LeaseDetails.UnitNumber = '" & szaUnit(0) & "';"
      
      rstLessee.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly
      
      sSQLQuery_ = "SELECT * FROM tlbLetterReports"
      
      rstReport.Open sSQLQuery_, conUnit_, adOpenDynamic, adLockOptimistic
      
      rstReport.AddNew
      rstReport!SageAccountNumber = rstLessee!SageAccountNumber
      rstReport!UnitNo = szaUnit(0)
      rstReport!SenderPosition = txtPosition.text
      rstReport!SenderName = txtSender.text
      rstReport!Body = txtBody.text
      rstReport!Subject = txtSubject.text
      rstReport!PrintDate = Format(Date, "dd/mm/yyyy")
      rstReport!LesseeName = IIf(IsNull(rstLessee!Name), "", rstLessee!Name)
      rstReport!AddressLine1 = IIf(IsNull(rstLessee!HOAddressLine1), "", rstLessee!HOAddressLine1)
      rstReport!AddressLine2 = IIf(IsNull(rstLessee!HOAddressLine2), "", rstLessee!HOAddressLine2)
      rstReport!AddressLine3 = IIf(IsNull(rstLessee!HOAddressLine3), "", rstLessee!HOAddressLine3)
      rstReport!AddressLine4 = IIf(IsNull(rstLessee!HOAddressLine4), "", rstLessee!HOAddressLine4)
      rstReport!PostCode = IIf(IsNull(rstLessee!HOPostCode), "", rstLessee!HOPostCode)
      rstReport.Update
      rstLessee.Close
      rstReport.Close
   Next i
   
   conUnit_.Close

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LetterTemplate.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   ' Passing the from and to date values to Crystal Reports
   Report.ParameterFields(1).AddCurrentValue CStr(txtSubject.text)
   Report.ParameterFields(2).AddCurrentValue txtBody.text
   Report.ParameterFields(3).AddCurrentValue CStr(txtSender.text)
   Report.ParameterFields(4).AddCurrentValue CStr(txtPosition.text)

   Load frmReport
   frmReport.LoadReportViewer Report

   Unload Me
End Sub

Private Sub cmdpropertyunit_Click()
'
'   Dim szSQL As String
'
'   If bSaved = False Then
'       If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Template") = vbYes Then
'         validateSave
'         If bValidate = False Then
'            Exit Sub
'         End If
'       End If
'   End If
'
'   adcProUnit.ConnectionString = getConnectionString
'   szSQL = "SELECT * FROM Property;"
'
'   adcProUnit.RecordSource = szSQL
'   adcProUnit.CommandType = adCmdText
'   adcProUnit.Refresh
'
'   lstProperty.Clear
'   While Not adcProUnit.Recordset.EOF
'      lstProperty.AddItem adcProUnit.Recordset.Fields.Item("PropertyID").Value & " \ " & _
'                          adcProUnit.Recordset.Fields.Item("PropertyName").Value
'      adcProUnit.Recordset.MoveNext
'   Wend
'
'   fraPropertyUnit.Top = 0
'   fraPropertyUnit.Left = 0
'   fraPropertyUnit.Visible = True
'   fraLetterBody.Visible = False
End Sub

Private Sub cmdRemUnit_Click()
   Dim j As Integer

   If lstDisplay.SelCount < 1 Then Exit Sub
   
   Do
      If lstDisplay.Selected(j) Then
         lstUnit.AddItem lstDisplay.text
         lstDisplay.RemoveItem lstDisplay.ListIndex
         j = j - 1
      End If
      j = j + 1
   Loop While j <= lstDisplay.ListCount - 1
End Sub

Private Sub cmdRemUnitAll_Click()
   If lstDisplay.ListCount < 0 Then Exit Sub
   lstDisplay.Clear
End Sub

Private Sub cmdSelectdiff_Click(Index As Integer)
   Dim sSQLQuery_ As String

   fraListTemplates.Visible = True
   fraLetterBody.Visible = False
'
'   sSQLQuery_ = "SELECT TemplateID, TemplateName, ModifiedBy, Description, ModifiedDate " & _
'                "FROM Template;"
'
'   adcTemplates.RecordSource = sSQLQuery_
'   adcTemplates.CommandType = adCmdText
'   adcTemplates.Refresh
End Sub

Private Sub Form_Activate()
   If szLetter = "LT" Then Me.Caption = "Letter Templates"
   If szLetter = "RT" Then Me.Caption = "Reminder Templates"

   Dim sSQLQuery_    As String
   Dim listSQL       As String

   fraListTemplates.Top = 0
   fraListTemplates.Left = 0

   fraLetterBody.Top = 0
   fraLetterBody.Left = 0

   ConfigFlxTemplates

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadFlxTemplate adoConn
      
   LoadTempType "TEMP_TYPE", adoConn
   
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_Load()
   Me.Height = 4710
   Me.Width = 11580
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   fraListTemplates.BackColor = Me.BackColor
   fraLetterBody.BackColor = Me.BackColor
   fraPropertyUnit.BackColor = Me.BackColor

   bChoice = False
   bSaved = False
   bValidate = True

   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadFlxTemplate(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim iRow    As Integer
   Dim adoRST  As New ADODB.Recordset

   szSQL = "SELECT TemplateID, TemplateName, ModifiedBy, Description, " & _
                  "TemplateDate, Salutation, Body, SenderName, SenderPosition, TempType, OurRef " & _
           "FROM Template " & _
           "WHERE TemplateName <> 'BACS Email Template' AND " & _
                 "TemplateName <> 'Demand Email Template' AND " & _
                 "TempType = '" & szLetter & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   With flxTemplates
      While Not adoRST.EOF
         .TextMatrix(iRow, 0) = adoRST.Fields.Item("TemplateID").Value
         .TextMatrix(iRow, 1) = adoRST.Fields.Item("TemplateName").Value
         .TextMatrix(iRow, 2) = adoRST.Fields.Item("ModifiedBy").Value
         .TextMatrix(iRow, 3) = IIf(IsNull(adoRST.Fields.Item("Description").Value), "", adoRST.Fields.Item("Description").Value) 'adoRST.Fields.Item("Description").Value
         .TextMatrix(iRow, 4) = Format(adoRST.Fields.Item("TemplateDate").Value, "dd/mm/yyyy")
         .TextMatrix(iRow, 5) = IIf(IsNull(adoRST.Fields.Item("Salutation").Value), "", adoRST.Fields.Item("Salutation").Value) 'adoRST.Fields.Item("Salutation").Value
         .TextMatrix(iRow, 6) = adoRST.Fields.Item("Body").Value
         .TextMatrix(iRow, 7) = IIf(IsNull(adoRST.Fields.Item("SenderName").Value), "", adoRST.Fields.Item("SenderName").Value) 'adoRST.Fields.Item("SenderName").Value
         .TextMatrix(iRow, 8) = IIf(IsNull(adoRST.Fields.Item("SenderPosition").Value), "", adoRST.Fields.Item("SenderPosition").Value) 'adoRST.Fields.Item("SenderPosition").Value
         .TextMatrix(iRow, 9) = adoRST.Fields.Item("TempType").Value
         .TextMatrix(iRow, 10) = IIf(IsNull(adoRST.Fields.Item("OurRef").Value), "", adoRST.Fields.Item("OurRef").Value)
         

         adoRST.MoveNext
         If Not adoRST.EOF Then .AddItem ""
         iRow = iRow + 1
      Wend
   End With
   
   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ConfigFlxTemplates()
   Dim szFlxHeader As String

   szFlxHeader$ = "|<|<|<|<||||||"

   With flxTemplates
      .FormatString = szFlxHeader$
      .Cols = 11
      .ColWidth(0) = 0
      .ColWidth(1) = 4200
      .ColWidth(2) = 1300
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      .ColWidth(9) = 0
      .ColWidth(10) = 0
      .Rows = 2
      .RowHeight(0) = 0
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub lstProperty_Click()
   Dim xSQL As String
   Dim ST() As String
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim rsLessee As New ADODB.Recordset
   ST = Split(lstProperty.text, " \ ")

   Label9.Caption = Label9.Caption & ST(1)

   xSQL = "SELECT UnitName,Units.UnitNumber FROM Units,LeaseDetails " & _
          "WHERE  Units.PropertyID = '" & ST(0) & "' and " & _
                 "Units.UnitNumber = LeaseDetails.UnitNumber And LeaseDetails.Status = True;"

'   adcProUnit.RecordSource = xSQL
'   adcProUnit.CommandType = adCmdText
'   adcProUnit.Refresh
   rsLessee.Open xSQL, adoConn, adOpenKeyset, adLockOptimistic
   lstUnit.Clear
  
'   While Not adcProUnit.Recordset.EOF
'      lstUnit.AddItem adcProUnit.Recordset.Fields.Item("UnitNumber").Value & " \ " & _
'                       adcProUnit.Recordset.Fields.Item("UnitName").Value
'      adcProUnit.Recordset.MoveNext
'   Wend
    While Not rsLessee.EOF
          lstUnit.AddItem rsLessee.Fields.Item("UnitNumber").Value & " \ " & _
                           rsLessee.Fields.Item("UnitName").Value
          rsLessee.MoveNext
      Wend
'   lstUnit.MultiSelect 2
    rsLessee.Close
    adoConn.Close
End Sub

Private Sub lstUnit_DblClick()
   cmdAddUnit_Click
End Sub

Private Sub txtBody_Change()
   bSaved = False
End Sub

Private Sub txtOurRef_Change()
    bSaved = False
End Sub

Private Sub txtPosition_Change()
   bSaved = False
End Sub

Private Sub txtSalutation_Change()
   bSaved = False
End Sub

Private Sub txtSender_Change()
   bSaved = False
End Sub

Private Sub txtSPDate_Change()
   TextBoxChangeDate txtSPDate
   bSaved = False
End Sub

Private Sub txtSPDate_GotFocus()
   SelTxtInCtrl txtSPDate
End Sub

Private Sub txtSPDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSPDate, KeyAscii
End Sub

Private Sub txtSPDate_LostFocus()
   TextBoxFormatDate txtSPDate
End Sub

Private Sub txtSubject_Change()
   bSaved = False
End Sub

Private Sub txtTemplateName_Change()
   bSaved = False
End Sub

Private Sub cmdSelect_Click()
   If flxTemplates.Rows < 1 Then
      ShowMsgInTaskBar "There is no saved template.", "Y", "N"
      Exit Sub
   End If
   If flxTemplates.TextMatrix(flxTemplates.row, 0) = "" Then
      ShowMsgInTaskBar "Please select a record to edit.", "Y", "N"
      Exit Sub
   End If

   fraListTemplates.Visible = False
   fraLetterBody.Visible = True

   With flxTemplates
      txtTemplateName.text = .TextMatrix(.row, 1)
      txtSPDate.text = .TextMatrix(.row, 4)
      txtSalutation.text = .TextMatrix(.row, 5)
      txtSubject.text = .TextMatrix(.row, 3)
      txtBody.text = .TextMatrix(.row, 6)
      txtSender.text = .TextMatrix(.row, 7)
      txtPosition.text = .TextMatrix(.row, 8)
      cmbTempType.Value = .TextMatrix(.row, 9)
      txtOurRef.text = .TextMatrix(.row, 10)
   End With

   bChoice = False
   bSaved = True
End Sub

Private Sub cmdSave_Click()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String
    If txtTemplateName.text = "" Then
      ShowMsgInTaskBar "Please type the template name.", "Y", "N"
      txtTemplateName.SetFocus
      Exit Sub
   Else
      If txtTemplateName.text = "BACS Email Template" Or _
               txtTemplateName.text = "Demand Email Template" Then
         ShowMsgInTaskBar "Not a valid template name.", "Y", "N"
         txtTemplateName.SetFocus
         Exit Sub
      End If
   End If
   If cmbTempType.text = "" Then
      ShowMsgInTaskBar "Please type the template type.", "Y", "N"
      cmbTempType.SetFocus
      Exit Sub
   End If
   If txtSPDate.text = "" Then
      ShowMsgInTaskBar "Please type the template date.", "Y", "N"
      txtSPDate.SetFocus
      Exit Sub
   End If
'   If txtSalutation.text = "" Then
'      ShowMsgInTaskBar "Please type the salutation.", "Y", "N"
'      txtSalutation.SetFocus
'      Exit Sub
'   End If
'  If txtOurRef.text = "" Then
'      ShowMsgInTaskBar "Please type the our reference.", "Y", "N"
'      txtOurRef.SetFocus
'      Exit Sub
'   End If
'   If txtSubject.text = "" Then
'      ShowMsgInTaskBar "Please type the subject.", "Y", "N"
'      txtSubject.SetFocus
'      Exit Sub
'   End If
'   If txtBody.text = "" Then
'      ShowMsgInTaskBar "Please type the template body.", "Y", "N"
'      txtBody.SetFocus
'      Exit Sub
'   End If
'   If txtSender.text = "" Then
'      ShowMsgInTaskBar "Please type the sender name.", "Y", "N"
'      txtSender.SetFocus
'      Exit Sub
'   End If
   
   
   adoConn.Open getConnectionString
   
   If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Template") = vbNo Then
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   Else
      validateSave adoConn

      ConfigFlxTemplates
      LoadFlxTemplate adoConn

      adoConn.Close
      Set adoConn = Nothing

      If bValidate = False Then
         Exit Sub
      End If
   End If

   fraListTemplates.Visible = True
   fraLetterBody.Visible = False
End Sub

Private Sub validateSave(adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim i          As Integer
   Dim existSame  As Boolean
   Dim adoRST     As New ADODB.Recordset
   
   bValidate = False
   

   If bChoice Then
      For i = 0 To flxTemplates.Rows - 1
         If flxTemplates.TextMatrix(i, 1) = txtTemplateName.text Then
            If MsgBox("Do you want to overwrite existing template?", vbQuestion + vbYesNo, "Template Exists") = vbYes Then
               flxTemplates.row = i
               i = flxTemplates.Rows - 1
            Else
               txtTemplateName.SetFocus
               Exit Sub
            End If
            Exit For
         End If
      Next i
   End If

   bValidate = True

   With adoRST
      '  SAVE data
      If bChoice Then            'Saving new
         szSQL = "SELECT TemplateName, SenderName, SenderPosition, Salutation, Body, " & _
                        "TemplateDate, Description, ModifiedBy, TempType,OurRef " & _
                 "FROM Template;"
         .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

         .AddNew
         .Fields.Item(0).Value = txtTemplateName.text
         .Fields.Item(1).Value = txtSender.text
         .Fields.Item(2).Value = txtPosition.text
         .Fields.Item(3).Value = txtSalutation.text
         .Fields.Item(4).Value = txtBody.text
         .Fields.Item(5).Value = Format(txtSPDate.text, "dd mmmm yyyy")
         .Fields.Item(6).Value = txtSubject.text
         .Fields.Item(7).Value = frmMMain.SystemUserName
         .Fields.Item(8).Value = cmbTempType.Value
         .Fields.Item(9).Value = txtOurRef.text
         .Update
      Else
         szSQL = "Select TemplateName, SenderName, SenderPosition, Salutation, Body, " & _
                        "TemplateDate, Description, ModifiedBy, TempType, OurRef  " & _
                 "FROM  Template " & _
                 "WHERE TemplateID = " & flxTemplates.TextMatrix(flxTemplates.row, 0) & ";"
         .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

         .Fields.Item(0).Value = txtTemplateName.text
         .Fields.Item(1).Value = txtSender.text
         .Fields.Item(2).Value = txtPosition.text
         .Fields.Item(3).Value = txtSalutation.text
         .Fields.Item(4).Value = txtBody.text
         .Fields.Item(5).Value = Format(txtSPDate.text, "dd mmmm yyyy")
         .Fields.Item(6).Value = txtSubject.text
         .Fields.Item(7).Value = frmMMain.SystemUserName
         .Fields.Item(8).Value = cmbTempType.Value
          .Fields.Item(9).Value = txtOurRef.text
         .Update
      End If
      .Close
   End With
   Set adoRST = Nothing

   bSaved = True

   ShowMsgInTaskBar "The Data is Succesfully Saved.", "Y", "P"
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
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
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
