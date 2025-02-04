VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmConfigEB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configure Electronic Banking"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigEB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Bank Account Details:"
      Height          =   975
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   9675
      Begin VB.Label lblSortCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Sort Code"
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblAccNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Num"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Code:"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblBankAccountName 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancelChange 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   6240
      TabIndex        =   13
      Top             =   3435
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditClient 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   1560
      TabIndex        =   11
      Top             =   3435
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveClient 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   3900
      TabIndex        =   12
      Top             =   3435
      Width           =   1215
   End
   Begin VB.Frame fraEntry 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2235
      Left            =   135
      TabIndex        =   15
      Top             =   1080
      Width           =   9600
      Begin VB.TextBox txtProcessedFileLocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1305
         Width           =   6540
      End
      Begin VB.TextBox txtBACSeMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1995
         TabIndex        =   20
         Top             =   1725
         Visible         =   0   'False
         Width           =   7050
      End
      Begin VB.TextBox txtFileIdentifire 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1995
         TabIndex        =   8
         Top             =   495
         Width           =   3150
      End
      Begin VB.TextBox txtFileExten 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7470
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "*.csv"
         Top             =   570
         Width           =   1575
      End
      Begin VB.TextBox txtEB_OP_FL 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   905
         Width           =   6540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process File Location "
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   24
         Top             =   1305
         Width           =   1515
      End
      Begin MSForms.CommandButton cmdProcessedFileLocation 
         Height          =   315
         Left            =   8640
         TabIndex        =   23
         ToolTipText     =   "Edit the Path"
         Top             =   1305
         Width           =   375
         Size            =   "661;556"
         Picture         =   "frmConfigEB.frx":030A
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BACS Email:"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   21
         Top             =   1725
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSForms.ComboBox cboEbanking 
         Height          =   300
         Left            =   1995
         TabIndex        =   6
         Top             =   120
         Width           =   7140
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "12594;529"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-banking service:"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File identifier:"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   18
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "File extension:"
         Height          =   255
         Index           =   5
         Left            =   6270
         TabIndex        =   17
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BACS File location:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   16
         Top             =   900
         Width           =   1320
      End
      Begin MSForms.CommandButton cmdPathEdit 
         Height          =   315
         Left            =   8640
         TabIndex        =   10
         ToolTipText     =   "Edit the Path"
         Top             =   900
         Width           =   375
         Size            =   "661;556"
         Picture         =   "frmConfigEB.frx":08A4
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "frmConfigEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public idBank As String

Private Sub cboClient_Change()

End Sub

Private Sub cboEbanking_Change()
    If cboEbanking.ListIndex = 0 Or cboEbanking.ListIndex = 2 Then
            Label1(8).Visible = False
            txtProcessedFileLocation.Visible = False
            cmdProcessedFileLocation.Visible = False
    Else
            Label1(8).Visible = True
            txtProcessedFileLocation.Visible = True
            cmdProcessedFileLocation.Visible = True
    End If
End Sub

Private Sub cmdCancelChange_Click()
   Unload Me
End Sub

Private Sub cmdEditClient_Click()
   fraEntry.Enabled = True
   cmdEditClient.Enabled = False
   cmdSaveClient.Enabled = True
End Sub

Private Sub cmdPathEdit_Click()
   Dim szResFolder As String

   szResFolder = BrowseForFolder(hWnd, "Please select a folder.")

   If szResFolder = "" Then
      MsgBox "You have not selected any folder path.", vbInformation + vbOKOnly, "Folder Path"

      Exit Sub
   Else
      txtEB_OP_FL.text = szResFolder
   End If
End Sub

Private Sub cmdProcessedFileLocation_Click()
   Dim szResFolder As String

   szResFolder = BrowseForFolder(hWnd, "Please select a folder.")

   If szResFolder = "" Then
      MsgBox "You have not selected any folder path.", vbInformation + vbOKOnly, "Folder Path"
      txtProcessedFileLocation.text = ""
      Exit Sub
   Else
      txtProcessedFileLocation.text = szResFolder
   End If
End Sub

Private Sub cmdSaveClient_Click()
   If (Len(txtFileExten.text) < 3 And Len(txtFileExten.text) > 0) Then
      MsgBox "Please enter the file extension.", vbCritical + vbOKOnly, "Electronic Banking"
      txtFileExten.SetFocus
      Exit Sub
   End If
   If txtEB_OP_FL.text = "" Then
      MsgBox "Please enter the file location.", vbCritical + vbOKOnly, "Electronic Banking"
      txtEB_OP_FL.SetFocus
      Exit Sub
   End If
   If txtFileIdentifire.text = "" Then
      MsgBox "Please enter the file indentifier.", vbCritical + vbOKOnly, "Electronic Banking"
      txtFileIdentifire.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

'Issue 832 warning for different output
   Dim rsDiffOutput As New ADODB.Recordset
   rsDiffOutput.Open "Select FileLoc from tlbclientBanks", adoConn, adOpenDynamic, adLockOptimistic
   If Not rsDiffOutput.EOF Then
        If Trim(txtEB_OP_FL.text) <> Trim(rsDiffOutput("FileLoc").Value) And cboEbanking.ListIndex = 1 Then
            MsgBox "The BACS output location selected " & txtEB_OP_FL.text & " is different from the existing BACS output locations selected for your other bank accounts. Please ensure that the BACS output location selected is the same for all your bank accounts.", vbInformation, "Warning"
            rsDiffOutput.Close
'            adoConn.Close
'            Exit Sub
        End If
   End If
   'rsDiffOutput.Close
   'Set rsDiffOutput = Nothing
   szSQL = "SELECT * " & _
           "FROM tlbClientBanks " & _
           "WHERE MY_ID = " & idBank & ";"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   adoRst.Fields.Item("EB").Value = cboEbanking.Column(0)
   adoRst.Fields.Item("Indentifier").Value = txtFileIdentifire.text
   adoRst.Fields.Item("FileExten").Value = txtFileExten.text
   adoRst.Fields.Item("FileLoc").Value = txtEB_OP_FL.text
   adoRst.Fields.Item("email").Value = txtBACSeMail.text
   'This line has been added by anol 2020-03-12
   adoRst.Fields.Item("ProcessFileLoc").Value = txtProcessedFileLocation.text

   adoRst.Update
   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   fraEntry.Enabled = False
   cmdEditClient.Enabled = True
   cmdSaveClient.Enabled = False
End Sub

Private Sub CommandButton1_Click()
    
End Sub

Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
'   Me.Height = 3735
'   Me.Width = 6270
   Me.BackColor = MODULEBACKCOLOR

   fraEntry.Enabled = False

   Call E_BankingService 'read \BACS\bacs.txt file and loads the first combo

   LoadEBDetails
End Sub

Public Sub E_BankingService()
   Dim szLine As String
   Dim Data() As String
   Dim i As Integer

   On Error GoTo CatchErr

   Open App.Path & "\BACS\bacs.txt" For Input As #1

   ReDim Data(1, 0) As String

   i = 0
   While Not EOF(1)
      Line Input #1, szLine
      If Val(szLine) > 0 Then
         Data(0, i) = szLine
         Line Input #1, szLine
         Data(1, i) = szLine
         Line Input #1, szLine
         i = i + 1
         ReDim Preserve Data(1, i) As String
      End If
   Wend
   cboEbanking.Column() = Data()

CatchErr:
   Close #1
End Sub

Private Sub LoadEBDetails()
   lblBankAccountName.Caption = frmClientNew4.txtBank_AC_Name.text
   lblAccNum.Caption = frmClientNew4.txtBANK_AC_NUM.text
   lblSortCode.Caption = frmClientNew4.txtBANK_SC.text

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks " & _
           "WHERE MY_ID = " & idBank & ";"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   cboEbanking.ListIndex = DropDownListPoint(cboEbanking, IIf(IsNull(adoRst.Fields.Item("EB").Value), "", adoRst.Fields.Item("EB").Value))
   txtFileIdentifire.text = IIf(IsNull(adoRst.Fields.Item("Indentifier").Value), Format(Now, "ddmmyyyy") & "_1", adoRst.Fields.Item("Indentifier").Value)
   txtFileExten.text = IIf(IsNull(adoRst.Fields.Item("FileExten").Value), "*.csv", adoRst.Fields.Item("FileExten").Value)
   txtEB_OP_FL.text = IIf(IsNull(adoRst.Fields.Item("FileLoc").Value), "", adoRst.Fields.Item("FileLoc").Value)
   txtBACSeMail.text = IIf(IsNull(adoRst.Fields.Item("email").Value), "", adoRst.Fields.Item("email").Value)
   
   txtProcessedFileLocation.text = IIf(IsNull(adoRst.Fields.Item("ProcessFileLoc").Value), "", adoRst.Fields.Item("ProcessFileLoc").Value)

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmClientNew4.Enabled = True
   Unload Me
End Sub

Private Sub txtFileExten_KeyPress(KeyAscii As Integer)
'   If Len(txtFileExten.text) = 0 Then Exit Sub
   If KeyAscii = 45 Or KeyAscii = 95 Or KeyAscii = 27 Or KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub

   If KeyAscii = 42 Then
      If InStr(txtFileExten.text, "*") > 0 Then
         KeyAscii = 0
         Exit Sub
      End If
   End If
   If KeyAscii = 46 Then
      If InStr(txtFileExten.text, ".") > 0 Then
         KeyAscii = 0
         Exit Sub
      End If
   End If
   
   If KeyAscii > 64 And KeyAscii < 91 Then
      KeyAscii = KeyAscii + 32
   End If
End Sub

Private Sub txtFileExten_LostFocus()
   If txtFileExten.text = "" Then
      MsgBox "Please enter the file extension.", vbCritical + vbOKOnly, "Electronic Banking"
      txtFileExten.SetFocus
      Exit Sub
   End If
   If Left(txtFileExten.text, 2) <> "*." Then
      txtFileExten.text = "*." & txtFileExten.text
   End If
End Sub
