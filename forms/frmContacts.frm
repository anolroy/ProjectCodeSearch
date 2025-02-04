VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmContacts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8280
   Begin VB.CheckBox Check4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Email Purchase Order"
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
      Left            =   5020
      TabIndex        =   41
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Email Remittance"
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
      Left            =   5020
      TabIndex        =   40
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Email Purchase Order"
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
      Left            =   1060
      TabIndex        =   39
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Email Remittance"
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
      Left            =   1060
      TabIndex        =   38
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton cmdUnitLookup 
      Caption         =   """"
      Height          =   285
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4200
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel Changes"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   1515
   End
   Begin VB.TextBox txtAddressLine4 
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
      Left            =   5040
      TabIndex        =   13
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtAddressLine2 
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
      Left            =   5040
      TabIndex        =   11
      Top             =   1035
      Width           =   2655
   End
   Begin VB.TextBox txtPostCode 
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
      Left            =   5040
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtAddressLine3 
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
      Left            =   5040
      TabIndex        =   12
      Top             =   1365
      Width           =   2655
   End
   Begin VB.TextBox txtAddressLine1 
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
      Left            =   5040
      TabIndex        =   10
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtPersonalEmail 
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
      Left            =   5040
      MaxLength       =   100
      TabIndex        =   15
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtMobile 
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
      Left            =   5040
      TabIndex        =   16
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtHomeTel 
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
      Left            =   5040
      TabIndex        =   17
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtOfficeAddressLine4 
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
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtContactName 
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
      Left            =   1100
      TabIndex        =   0
      Top             =   120
      Width           =   2865
   End
   Begin VB.TextBox txtOfficeAddressLine2 
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
      Left            =   1080
      TabIndex        =   4
      Top             =   1035
      Width           =   2655
   End
   Begin VB.TextBox txtOfficePostCode 
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtOfficeAddressLine3 
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
      Left            =   1080
      TabIndex        =   23
      Top             =   1365
      Width           =   2655
   End
   Begin VB.TextBox txtOfficeAddressLine1 
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
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtOfficeEmail 
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
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   7
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtOfficeMobile 
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
      Left            =   1080
      TabIndex        =   8
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtOfficeTel 
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
      Left            =   1080
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4683;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4155
      TabIndex        =   37
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   4200
      TabIndex        =   36
      Top             =   3240
      Width           =   255
   End
   Begin MSForms.CommandButton cmdSendMail 
      Height          =   285
      Index           =   0
      Left            =   4605
      TabIndex        =   22
      Top             =   2520
      Width           =   420
      Size            =   "741;503"
      Picture         =   "frmContacts.frx":5F8A
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      Caption         =   " Alternative Address:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4260
      TabIndex        =   35
      Top             =   480
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4155
      TabIndex        =   34
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4155
      TabIndex        =   33
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   4155
      TabIndex        =   32
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4155
      TabIndex        =   31
      Top             =   2880
      Width           =   525
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      Caption         =   " Main Address:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   30
      Top             =   480
      Width           =   1020
   End
   Begin MSForms.CommandButton cmdSendMail 
      Height          =   285
      Index           =   1
      Left            =   645
      TabIndex        =   21
      Top             =   2520
      Width           =   420
      Size            =   "741;503"
      Picture         =   "frmContacts.frx":BBAC
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   200
      TabIndex        =   29
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   195
      TabIndex        =   28
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   195
      TabIndex        =   27
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   195
      TabIndex        =   26
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   195
      TabIndex        =   25
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   195
      TabIndex        =   24
      Top             =   2880
      Width           =   525
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   3585
      Left            =   120
      Top             =   585
      Width           =   3825
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   3585
      Left            =   4080
      Top             =   585
      Width           =   3825
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WHOS_CONTACT  As String
Public LOADED        As Boolean
Public LOADING_MODE  As String
Public LOADING_ID    As String

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
   If txtContactName.text = "" Then
      ShowMsgInTaskBar "Please enter the contact name", "Y", "N"
      txtContactName.SetFocus

      Exit Sub
   End If

   Dim adoConn       As New ADODB.Connection
   Dim adoRST        As New ADODB.Recordset
   Dim szSQL         As String

   If LOADING_MODE = "NEW" Then
      szSQL = "SELECT * FROM Contacts;"
   Else
      szSQL = "SELECT * FROM Contacts WHERE ContactID = '" & LOADING_ID & "';"
   End If
   adoConn.Open getConnectionString

   With adoRST
      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
      If LOADING_MODE = "NEW" Then
         .AddNew
         .Fields.Item("ContactID").Value = UniqueID()
         .Fields.Item("WhosContact").Value = WHOS_CONTACT
      End If
      
      If WHOS_CONTACT = "S" Then _
         .Fields.Item("HeadContact").Value = frmManagingAgent3.txtMAID.text      'Managing agent
      If WHOS_CONTACT = "C" Then _
         .Fields.Item("HeadContact").Value = WHOS_CONTACT      'Client
      If WHOS_CONTACT = "T" Then _
         .Fields.Item("HeadContact").Value = WHOS_CONTACT      'Tenant
      If WHOS_CONTACT = "M" Then _
         .Fields.Item("HeadContact").Value = WHOS_CONTACT      'Managing Agent

      .Fields.Item("ContactName").Value = txtContactName.text
      .Fields.Item("AddressLine1").Value = txtAddressLine1.text
      .Fields.Item("AddressLine2").Value = txtAddressLine2.text
      .Fields.Item("AddressLine3").Value = txtAddressLine3.text
      .Fields.Item("AddressLine4").Value = txtAddressLine4.text
      .Fields.Item("PostCode").Value = txtPostCode.text
      .Fields.Item("PersonalEmail").Value = txtPersonalEmail.text
      .Fields.Item("HomeTel").Value = txtHomeTel.text
      .Fields.Item("Mobile").Value = txtMobile.text
      .Fields.Item("OfficeAddressLine1").Value = txtOfficeAddressLine1.text
      .Fields.Item("OfficeAddressLine2").Value = txtOfficeAddressLine2.text
      .Fields.Item("OfficeAddressLine3").Value = txtOfficeAddressLine3.text
      .Fields.Item("OfficeAddressLine4").Value = txtOfficeAddressLine4.text
      .Fields.Item("OfficePostCode").Value = txtOfficePostCode.text
      .Fields.Item("OfficeTel").Value = txtOfficeTel.text
      .Fields.Item("OfficeEmail").Value = txtOfficeEmail.text
      .Fields.Item("OfficeMobile").Value = txtOfficeMobile.text

      .Update
      .Close
   End With

   Set adoRST = Nothing

   If WHOS_CONTACT = "S" Then _
      frmManagingAgent3.LoadFlxContact adoConn
      frmManagingAgent3.Enabled = True
   adoConn.Close
   Set adoConn = Nothing

   ClearForm
End Sub

Private Sub ClearForm()
   Dim ctrlControl As Control

   For Each ctrlControl In Me.Controls
      If TypeName(ctrlControl) = "TextBox" Then _
         ctrlControl.text = ""
   Next ctrlControl
End Sub

Private Sub Form_Activate()
   If LOADED Then Exit Sub
   
   If LOADING_MODE = "EDIT" Then
      Dim adoConn       As New ADODB.Connection
      Dim adoRST        As New ADODB.Recordset
      Dim szSQL         As String
      
      szSQL = "SELECT * FROM Contacts WHERE ContactID = '" & LOADING_ID & "';"

      adoConn.Open getConnectionString

      With adoRST
         .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

         If Not .EOF Then
            txtContactName.text = .Fields.Item("ContactName").Value
            txtAddressLine1.text = .Fields.Item("AddressLine1").Value
            txtAddressLine2.text = .Fields.Item("AddressLine2").Value
            txtAddressLine3.text = .Fields.Item("AddressLine3").Value
            txtAddressLine4.text = .Fields.Item("AddressLine4").Value
            txtPostCode.text = .Fields.Item("PostCode").Value
            txtPersonalEmail.text = .Fields.Item("PersonalEmail").Value
            txtHomeTel.text = .Fields.Item("HomeTel").Value
            txtMobile.text = .Fields.Item("Mobile").Value
            txtOfficeAddressLine1.text = .Fields.Item("OfficeAddressLine1").Value
            txtOfficeAddressLine2.text = .Fields.Item("OfficeAddressLine2").Value
            txtOfficeAddressLine3.text = .Fields.Item("OfficeAddressLine3").Value
            txtOfficeAddressLine4.text = .Fields.Item("OfficeAddressLine4").Value
            txtOfficePostCode.text = .Fields.Item("OfficePostCode").Value
            txtOfficeTel.text = .Fields.Item("OfficeTel").Value
            txtOfficeEmail.text = .Fields.Item("OfficeEmail").Value
            txtOfficeMobile.text = .Fields.Item("OfficeMobile").Value

            .Update
            .Close
         End If
      End With

      Set adoRST = Nothing

      adoConn.Close
      Set adoConn = Nothing
   End If

   LOADED = True
End Sub

Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Width = 8115
   Me.Height = 5175
   LOADED = False

   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If WHOS_CONTACT = "S" Then
      frmSupplier.Enabled = True
   End If
End Sub
