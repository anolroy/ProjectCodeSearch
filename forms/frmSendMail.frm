VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSendMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Email"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12675
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "Send Mail"
      Height          =   975
      Left            =   5520
      Picture         =   "frmSendMail.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   975
      Left            =   10920
      Picture         =   "frmSendMail.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtxBody 
      Height          =   4335
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
      _Version        =   393217
      TextRTF         =   $"frmSendMail.frx":65D6
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   11055
   End
   Begin VB.Label lblRecipientAddress 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail Recipient:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSendMail_Click()
   If txtSubject.text = "" Then If MsgBox("Do you wish to send email without subject?", vbQuestion + vbYesNo, "Email") = vbNo Then Exit Sub

   Dim bEmailResult     As Boolean

   bEmailResult = SendEmail(szFromEmail, lblRecipientAddress.Caption, _
                           txtSubject.text, _
                           rtxBody.text, , , _
                           , frmLeasee1.txtTenantID.text, "Just Email")
   If bEmailResult Then
      ShowMsgInTaskBar "Email sent.", "Y", "P"

      SavingEmailInformation
      frmLeasee1.LoadFlxEmails

      Unload Me
   Else
      ShowMsgInTaskBar "No email sent.", "Y", "N"
   End If
End Sub

Private Sub SavingEmailInformation()
   Dim szLine  As String

'  even the file does not exists, system will create the file and start adding text at the bottom of the file
   Open DB_PATH & "\AllStuff\Logs\Email_" & SCID & "_" & frmLeasee1.txtTenantID.text & ".dat" For Append As #1

   szLine = "Email sent on:" & Format(Now, "dd/mm/yyyy") & "#" & Format(Now, "hh:nn") & "#" & UniqueID() & Chr$(13) & Chr$(10)
   szLine = szLine + "Email Address:" & lblRecipientAddress.Caption & Chr$(13) & Chr$(10)
   szLine = szLine + "Email Subject:" + txtSubject.text & Chr$(13) & Chr$(10)
   szLine = szLine + rtxBody.text
   szLine = szLine + "*****"

   Print #1, szLine
   Close #1
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmLeasee1.Enabled = True
   Unload Me
End Sub
