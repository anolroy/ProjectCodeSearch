VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEmailTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email Template"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailTemplate.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   13005
   Begin VB.Frame fraLetterBody 
      BorderStyle     =   0  'None
      Caption         =   "fraListTemplates"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         TabIndex        =   6
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   5
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Width           =   10575
      End
      Begin VB.TextBox txtBody 
         DataSource      =   "Adodc1"
         Height          =   3015
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   10575
      End
      Begin MSForms.ComboBox cboKeyWords 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   3720
         Width           =   2640
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4657;556"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;35277"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insert key words:"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   3720
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblBody 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Body"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmEmailTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bSave     As Boolean
Private bLoaded   As Boolean
Private szFocused As String
Private iTextPos  As Integer
Private mctlLast  As Control

Private Sub cboKeyWords_Click()
   If szFocused <> "" Then
      If mctlLast.text <> "" Then
         mctlLast.text = Left(mctlLast.text, iTextPos) & _
                        "<" & cboKeyWords.text & ">" & _
                        Right(mctlLast.text, Len(mctlLast.text) - iTextPos)
      Else
         mctlLast.text = "<" & cboKeyWords.text & ">"
      End If
   
   
      mctlLast.SetFocus
   End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdClose_GotFocus()
   szFocused = ""
End Sub

Private Sub cmdSave_Click()
   If Not bSave Then Exit Sub
   If txtSubject.text = "" Then Exit Sub
   If txtBody.text = "" Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String

   adoConn.Open getConnectionString

   If Me.Caption = "BACS Email Template" Then
      szSQL = "SELECT TemplateName, TemplateID, Description AS SUBJECT, Body, TemplateDate, ModifiedDate " & _
              "FROM Template " & _
              "WHERE TemplateName = 'BACS Email Template';"
   End If
   If Me.Caption = "Demand Email Template" Then
      szSQL = "SELECT TemplateName, TemplateID, Description AS SUBJECT, Body, TemplateDate, ModifiedDate " & _
              "FROM Template " & _
              "WHERE TemplateName = 'Demand Email Template';"
   End If
   If Me.Caption = "Statement Email Template" Then
      szSQL = "SELECT TemplateName, TemplateID, Description AS SUBJECT, Body, TemplateDate, ModifiedDate " & _
              "FROM Template " & _
              "WHERE TemplateName = 'Statement Email Template';"
   End If
   If Me.Caption = "Client Statement Email Template" Then
      szSQL = "SELECT TemplateName, TemplateID, Description AS SUBJECT, Body, TemplateDate, ModifiedDate " & _
              "FROM Template " & _
              "WHERE TemplateName = 'Client Statement Email Template';"
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   If adoRst.EOF Then adoRst.AddNew

   adoRst.Fields.Item("TemplateName").Value = Me.Caption
   adoRst.Fields.Item("SUBJECT").Value = txtSubject.text
   adoRst.Fields.Item("Body").Value = txtBody.text
   adoRst.Fields.Item("TemplateDate").Value = Format(Now, "dd mmmm yyyy")
   adoRst.Fields.Item("ModifiedDate").Value = Format(Now, "dd mmmm yyyy")
   adoRst.Update

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Template has been saved.", "Y", "P"
End Sub

Private Sub cmdSave_GotFocus()
   szFocused = ""
End Sub

Private Sub Form_Activate()
   If Not bLoaded Then
'      dim
      LoadExistingBody
      
      InsertCodes
      
      bLoaded = True
   End If
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 4710
   Me.Width = 11580
   Me.BackColor = MODULEBACKCOLOR
   fraLetterBody.BackColor = MODULEBACKCOLOR

   fraLetterBody.Top = 0
   fraLetterBody.Left = 0

   bSave = False
   bLoaded = False

   InsertCodes
End Sub

Public Sub InsertCodes()
   Dim adoConn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim SQLStr     As String
   Dim Data()     As String
   Dim i          As Integer
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer

   adoConn.Open getConnectionString

   SQLStr = "SELECT Code, Value " & _
            "FROM SecondaryCode " & _
            "WHERE SecondaryCode.PrimaryCode = 'DTKW' " & _
            "ORDER BY CODE;"
   adoRst.Open SQLStr, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   i = 0
   If Not adoRst.EOF Then
      While Not adoRst.EOF
         ReDim Preserve Data(1, i) As String
         Data(0, i) = CStr(adoRst!Code)
         Data(1, i) = CStr(adoRst!Value)
         adoRst.MoveNext
         i = i + 1
      Wend
   End If
   adoRst.Close
   Set adoRst = Nothing

   adoConn.Close
   Set adoConn = Nothing

   cboKeyWords.Clear
   cboKeyWords.Column() = Data()
End Sub

Private Sub LoadExistingBody()
   Dim adoConn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String

   adoConn.Open getConnectionString

   If Me.Caption = "BACS Email Template" Then
      szSQL = "SELECT TemplateID, Description AS SUBJECT, Body " & _
              "FROM Template " & _
              "WHERE TemplateName = 'BACS Email Template';"
   End If
   If Me.Caption = "Demand Email Template" Then
      szSQL = "SELECT TemplateID, Description AS SUBJECT, Body " & _
              "FROM Template " & _
              "WHERE TemplateName = 'Demand Email Template';"
   End If
   If Me.Caption = "Statement Email Template" Then
      szSQL = "SELECT TemplateID, Description AS SUBJECT, Body " & _
              "FROM Template " & _
              "WHERE TemplateName = 'Statement Email Template';"
   End If
   If Me.Caption = "Client Statement Email Template" Then
      szSQL = "SELECT TemplateID, Description AS SUBJECT, Body " & _
              "FROM Template " & _
              "WHERE TemplateName = 'Client Statement Email Template';"
   End If

'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      txtSubject.text = adoRst.Fields.Item("SUBJECT").Value
      txtBody.text = adoRst.Fields.Item("Body").Value
   Else
        If Me.Caption = "BACS Email Template" Or Me.Caption = "Demand Email Template" Then
            txtSubject.text = "Payment notification from <CLIENT NAME>"
            txtBody.text = "Payment notification from <CLIENT NAME>" & (Chr(13) + Chr(10)) & _
                           "Please see the attached remittance advice for details."
        Else
            'Salia did not gave me any specification
        End If
   End If

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub txtBody_Change()
   bSave = True
End Sub

Private Sub txtBody_GotFocus()
   szFocused = "Body"
End Sub

Private Sub txtBody_LostFocus()
   iTextPos = txtBody.SelStart
   Set mctlLast = txtBody
End Sub

Private Sub txtSubject_Change()
   bSave = True
End Sub

Private Sub txtSubject_GotFocus()
   szFocused = "Subject"
End Sub

Private Sub txtSubject_LostFocus()
   iTextPos = txtSubject.SelStart
   Set mctlLast = txtSubject
End Sub
