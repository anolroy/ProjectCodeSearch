VERSION 5.00
Begin VB.Form frmShoppingCentre 
   BackColor       =   &H00FFEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Setup"
   ClientHeight    =   5730
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Setup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8565
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFEEEE&
      Caption         =   "Second Contact Person"
      Height          =   1335
      Left            =   4560
      TabIndex        =   31
      Top             =   1320
      Width           =   3855
      Begin VB.TextBox txtCon2Name 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   645
         MaxLength       =   40
         TabIndex        =   34
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtCon2Mail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   645
         MaxLength       =   40
         TabIndex        =   33
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtCon2Phone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   645
         MaxLength       =   20
         TabIndex        =   32
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   105
         TabIndex        =   37
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "e-Mail:"
         Height          =   195
         Left            =   75
         TabIndex        =   36
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   195
         Left            =   60
         TabIndex        =   35
         Top             =   960
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFEEEE&
      Caption         =   "First Contact Person"
      Height          =   1335
      Left            =   4560
      TabIndex        =   24
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtCon1Name 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   660
         MaxLength       =   40
         TabIndex        =   27
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtCon1Phone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   660
         MaxLength       =   20
         ScrollBars      =   1  'Horizontal
         TabIndex        =   26
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtCon1Mail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   660
         MaxLength       =   40
         TabIndex        =   25
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "e-Mail:"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   195
         Left            =   75
         TabIndex        =   28
         Top             =   960
         Width           =   480
      End
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   23
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtDataSource 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtReference 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   9
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox txtCompanyWeb 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   8
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   4560
      MaxLength       =   200
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3000
      Width           =   3855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Data"
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtCompanyFax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   7
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox txtCompanyTel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   6
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtPC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtAdd4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtAdd3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtAdd2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtAdd1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel Changes"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   375
      Left            =   4020
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Datasource: (3rd Party)"
      Height          =   435
      Left            =   120
      TabIndex        =   22
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reference:"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   750
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   630
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   195
      Left            =   4560
      TabIndex        =   19
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   285
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frmShoppingCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim SQLStr As String
'
'Private Sub cboShopCentreName_LostFocus()
'   Dim i, j, match As Integer
'   match = 0
'
'   If cboShopCentreName.text <> "" Then
'       j = cboShopCentreName.ListCount - 1
'       For i = 0 To j
'           If cboShopCentreName.List(i) = cboShopCentreName.text Then
'               match = 1
'               Exit For
'           End If
'       Next i
'       If match = 0 Then
'           MsgBox "Company selected is invalid.", vbOKOnly + vbCritical, "Invalid Company"
'           cboShopCentreName.text = ""
'           Exit Sub
'       End If
'   End If
'End Sub

Private Sub cmdCancel_Click()
'   cboShopCentreName.Clear
   Call GetData
   Call DisableBoxes
End Sub

Private Sub cmdEdit_Click()
   Call EnableBoxes
End Sub

Private Sub cmdSave_Click()
   Dim temp As String
   Dim i As Integer

   If txtCompany.text = "" Then
       ShowMsgInTaskBar "You must enter a Shopping Centre Name.", , "N"
       Exit Sub
   End If
'   If cboShopCentreName.text = "" Then
'       MsgBox "You must select a Company from Sage.", vbOKOnly + vbCritical, "Missing Data"
'       Exit Sub
'   End If
'   If txtDataSource.text = "" Then
'       MsgBox "You must enter a Sage Datasource.", vbOKOnly + vbCritical, "Missing Data"
'       Exit Sub
'   End If

   'save new details to database
   Conn.Open getConnectionString

   rst.Open "SELECT * FROM ShoppingCentre", Conn, adOpenDynamic, adLockOptimistic

   If rst.RecordCount = 0 Then
       rst.AddNew
   Else
       rst.MoveFirst
   End If
'   For i = 2 To 3
'       If Mid(cboShopCentreName, i, 3) = " / " Then Rst!Code = Left(cboShopCentreName.text, i - 1)
'   Next i

   rst!SageDSN = IIf(txtDataSource.text = "", Null, txtDataSource.text)
   If txtCon1Name.text <> "" Then rst!Contact1 = txtCon1Name.text
   rst!Name = txtCompany.text
   If txtCon1Mail.text <> "" Then rst!Email1 = txtCon1Mail.text
   If txtCon1Phone.text <> "" Then rst!DirectLine1 = txtCon1Phone.text
   If txtCon2Name.text <> "" Then rst!Contact2 = txtCon2Name.text
   If txtCon2Mail.text <> "" Then rst!Email2 = txtCon2Mail.text
   If txtCon2Phone.text <> "" Then rst!DirectLine2 = txtCon2Phone.text
   If txtAdd1.text <> "" Then rst!AddressLine1 = txtAdd1.text
   If txtAdd2.text <> "" Then rst!AddressLine2 = txtAdd2.text
   If txtAdd3.text <> "" Then rst!AddressLine3 = txtAdd3.text
   If txtAdd4.text <> "" Then rst!AddressLine4 = txtAdd4.text
   If txtPC.text <> "" Then rst!PostCode = txtPC.text
   If txtCompanyTel.text <> "" Then rst!Telephone = txtCompanyTel.text
   If txtCompanyFax.text <> "" Then rst!Fax = txtCompanyFax.text
   If txtCompanyWeb.text <> "" Then rst!Website = txtCompanyWeb.text
   If txtReference.text <> "" Then rst!Reference = txtReference.text
   If txtNotes.text <> "" Then rst!Notes = txtNotes.text

   rst.Update
   rst.Close
   Conn.Close

   Conn.Open "DSN=PrestigeBMControlNS;UID=;PWD="

   Conn.Execute "UPDATE Databases SET SCName = '" & txtCompany.text & "' WHERE AccessDSN = '" & Adsn & "';"
'   Rst.Open "SELECT * FROM Databases WHERE AccessDSN = '" & Adsn & "'", Conn, adOpenDynamic, adLockOptimistic

'   Rst!SCName = txtCompany.text
'   Rst.Update
'   Rst.Close
   Conn.Close

   Call DisableBoxes
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   Me.Caption = "Property Details"
   Frame1.BackColor = Me.BackColor
   Frame2.BackColor = Me.BackColor

   Call DisableBoxes

   Call GetData
End Sub

Public Sub DisableBoxes()
'   cboShopCentreName.Enabled = False
   txtCompany.Enabled = False
   txtCon1Name.Enabled = False
   txtCon1Mail.Enabled = False
   txtCon1Phone.Enabled = False
   txtCon2Name.Enabled = False
   txtCon2Mail.Enabled = False
   txtCon2Phone.Enabled = False
   txtAdd1.Enabled = False
   txtAdd2.Enabled = False
   txtAdd3.Enabled = False
   txtAdd4.Enabled = False
   txtPC.Enabled = False
   txtCompanyTel.Enabled = False
   txtCompanyFax.Enabled = False
   txtNotes.Enabled = False
   txtCompanyWeb.Enabled = False
   txtReference.Enabled = False
   txtDataSource.Enabled = False

   cmdSave.Visible = False
   cmdCancel.Visible = False
   cmdEdit.Visible = True
End Sub

Public Sub EnableBoxes()
'   cboShopCentreName.Enabled = True
   txtCompany.Enabled = True
   txtCon1Name.Enabled = True
   txtCon1Mail.Enabled = True
   txtCon1Phone.Enabled = True
   txtCon2Name.Enabled = True
   txtCon2Mail.Enabled = True
   txtCon2Phone.Enabled = True
   txtAdd1.Enabled = True
   txtAdd2.Enabled = True
   txtAdd3.Enabled = True
   txtAdd4.Enabled = True
   txtPC.Enabled = True
   txtCompanyTel.Enabled = True
   txtCompanyFax.Enabled = True
   txtNotes.Enabled = True
   txtCompanyWeb.Enabled = True
   txtReference.Enabled = True
   txtDataSource.Enabled = True

   cmdSave.Visible = True
   cmdCancel.Visible = True
   cmdEdit.Visible = False
End Sub

Public Sub GetData()
   On Error GoTo ErrHandler

   Dim i As Integer

   Conn.Open getConnectionString

   SQLStr = "SELECT * FROM ShoppingCentre "
   rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If rst.RecordCount = 0 Then
       rst.Close
       Conn.Close
       Exit Sub
   Else
       txtDataSource.text = IIf(IsNull(rst!SageDSN), "", rst!SageDSN)
       txtCompany.text = gCurrentShopCentreName
       If IsNull(rst!Contact1) = False Then txtCon1Name.text = rst!Contact1
       If IsNull(rst!Email1) Then txtCon1Mail.text = "" Else txtCon1Mail.text = rst!Email1
       If IsNull(rst!DirectLine1) Then txtCon1Phone.text = "" Else txtCon1Phone.text = rst!DirectLine1
       If IsNull(rst!Contact2) Then txtCon2Name.text = "" Else txtCon2Name.text = rst!Contact2
       If IsNull(rst!Email2) Then txtCon2Mail.text = "" Else txtCon2Mail.text = rst!Email2
       If IsNull(rst!DirectLine2) Then txtCon2Phone.text = "" Else txtCon2Phone.text = rst!DirectLine2
       If IsNull(rst!AddressLine1) = False Then txtAdd1.text = rst!AddressLine1
       If IsNull(rst!AddressLine2) Then txtAdd2.text = "" Else txtAdd2.text = rst!AddressLine2
       If IsNull(rst!AddressLine3) Then txtAdd3.text = "" Else txtAdd3.text = rst!AddressLine3
       If IsNull(rst!AddressLine4) Then txtAdd4.text = "" Else txtAdd4.text = rst!AddressLine4
       If IsNull(rst!PostCode) Then txtPC.text = "" Else txtPC.text = rst!PostCode
       If IsNull(rst!Telephone) = False Then txtCompanyTel.text = rst!Telephone
       If IsNull(rst!Fax) Then txtCompanyFax.text = "" Else txtCompanyFax.text = rst!Fax
       If IsNull(rst!Website) Then txtCompanyWeb.text = "" Else txtCompanyWeb.text = rst!Website
       If IsNull(rst!Reference) Then txtReference.text = "" Else txtReference.text = rst!Reference
       If IsNull(rst!Notes) Then txtNotes.text = "" Else txtNotes.text = rst!Notes

       rst.Close
       Conn.Close

   End If

   Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        ShowMsgInTaskBar Err.Number & " - " & Err.description & " - " & Err.Source, , "N"
    End If
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub
'
'Private Sub mnuEdit_Click()
'   Call EnableBoxes
'End Sub
'
'Private Sub mnuExit_Click()
'   Unload frmMMain
'End Sub
'
'Private Sub mnuGlobal_Click()
'   Load frmGlobal
'   Unload Me
'   frmGlobal.Show
'End Sub

Private Sub txtcompany_LostFocus()

gCurrentShopCentreName = txtCompany.text
'Me.Caption = "Shopping Centre Details"

End Sub
