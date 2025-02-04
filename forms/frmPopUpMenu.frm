VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPopUpMenu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7575
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPopUpMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPurchaseOrder 
      Height          =   2895
      Left            =   240
      TabIndex        =   52
      Top             =   4320
      Width           =   3015
      Begin VB.CommandButton cmdPO_Posting 
         Caption         =   "&Post to History"
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
         Left            =   120
         TabIndex        =   57
         Top             =   2520
         Width           =   2775
      End
      Begin VB.OptionButton optPONoRange 
         Caption         =   "Document Reference Range"
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
         Left            =   165
         TabIndex        =   56
         Top             =   1580
         Width           =   2295
      End
      Begin VB.OptionButton optPODtRange 
         Caption         =   "Date Range"
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
         Left            =   160
         TabIndex        =   55
         Top             =   400
         Width           =   1140
      End
      Begin VB.OptionButton optSelPO 
         Caption         =   "Selected Purchase Order"
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
         Left            =   160
         TabIndex        =   54
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2720
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   63
         Top             =   1605
         Width           =   2775
         Begin VB.TextBox txtPORangeFrom 
            Appearance      =   0  'Flat
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
            Left            =   540
            TabIndex        =   65
            Top             =   375
            Width           =   855
         End
         Begin VB.TextBox txtPORangeTo 
            Appearance      =   0  'Flat
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
            Left            =   1800
            TabIndex        =   64
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "From:"
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
            Index           =   11
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "To:"
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
            Index           =   10
            Left            =   1500
            TabIndex        =   66
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   58
         Top             =   440
         Width           =   2775
         Begin VB.TextBox txtDtRangeToPO 
            Appearance      =   0  'Flat
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
            Left            =   960
            TabIndex        =   60
            Top             =   615
            Width           =   1215
         End
         Begin VB.TextBox txtDtRangeFromPO 
            Appearance      =   0  'Flat
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
            Left            =   960
            TabIndex        =   59
            Top             =   260
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "To:"
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
            Index           =   9
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "From:"
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
            Index           =   8
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame fraPurchasePosting 
      Height          =   2895
      Left            =   6120
      TabIndex        =   32
      Top             =   4080
      Width           =   3015
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2720
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton optSelPI 
         Caption         =   "Selected Invoices && Cr Notes"
         Height          =   255
         Left            =   160
         TabIndex        =   42
         Top             =   120
         Width           =   2415
      End
      Begin VB.OptionButton optPIDtRange 
         Caption         =   "Date Range"
         Height          =   255
         Left            =   160
         TabIndex        =   36
         Top             =   380
         Width           =   1215
      End
      Begin VB.OptionButton optFP_PI 
         Caption         =   "Fully Paid Invoices && Credit Notes"
         Height          =   255
         Left            =   160
         TabIndex        =   35
         Top             =   1440
         Width           =   2775
      End
      Begin VB.OptionButton optSlNoRange 
         Caption         =   "Document Reference Range"
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
         Left            =   165
         TabIndex        =   34
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmdPI_Posting 
         Caption         =   "&Post to History"
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Frame fraPIDtRange 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   37
         Top             =   400
         Width           =   2775
         Begin VB.TextBox txtDtRangeFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   38
            Top             =   260
            Width           =   1215
         End
         Begin VB.TextBox txtDtRangeTo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   39
            Top             =   615
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "From:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "To:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame fraPlRange 
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   43
         Top             =   1845
         Width           =   2775
         Begin VB.TextBox txtPlRangeTo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtPlRangeFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   540
            TabIndex        =   44
            Top             =   260
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "To:"
            Height          =   255
            Index           =   7
            Left            =   1500
            TabIndex        =   47
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "From:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraDemandPosting 
      Height          =   3735
      Left            =   3480
      TabIndex        =   14
      Top             =   3720
      Width           =   2535
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2240
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton optDPEmailDmds 
         Caption         =   "Emailed Demands"
         Height          =   255
         Left            =   160
         TabIndex        =   21
         Top             =   2000
         Width           =   2175
      End
      Begin VB.CommandButton cmdDP_Posting 
         Caption         =   "&Post to History"
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   2295
      End
      Begin VB.OptionButton optDPSlNoRange 
         Caption         =   "Demand No Range"
         Height          =   255
         Left            =   160
         TabIndex        =   22
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton optDPPrintedDmds 
         Caption         =   "Printed Demands"
         Height          =   255
         Left            =   160
         TabIndex        =   20
         Top             =   1720
         Width           =   2175
      End
      Begin VB.OptionButton optDPFPDemands 
         Caption         =   "Fully Paid Demands"
         Height          =   255
         Left            =   160
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton optDPDtRange 
         Caption         =   "Date Range"
         Height          =   255
         Left            =   160
         TabIndex        =   16
         Top             =   380
         Width           =   1215
      End
      Begin VB.Frame fraDPDtRange 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   400
         Width           =   2295
         Begin VB.TextBox txtDPDtRangeTo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   18
            Top             =   615
            Width           =   1215
         End
         Begin VB.TextBox txtDPDtRangeFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   17
            Top             =   260
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "To:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "From:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.OptionButton optDPSelDmds 
         Caption         =   "Selected Demands"
         Height          =   255
         Left            =   160
         TabIndex        =   15
         Top             =   120
         Width           =   2055
      End
      Begin VB.Frame fraDPSlRange 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   29
         Top             =   2325
         Width           =   2295
         Begin VB.OptionButton optSC 
            Caption         =   "SC"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton optSI 
            Caption         =   "SI"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtDPSlRangeFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            Top             =   260
            Width           =   855
         End
         Begin VB.TextBox txtDPSlRangeTo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   24
            Top             =   615
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "From:"
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   31
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "To:"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   30
            Top             =   600
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraBankCopy 
      Height          =   660
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1790
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   1760
         Y1              =   360
         Y2              =   360
      End
      Begin MSForms.Label lblBankCopy 
         Height          =   255
         Left            =   45
         TabIndex        =   9
         Top             =   120
         Width           =   1695
         BackColor       =   14737632
         Caption         =   "Copy Transaction"
         Size            =   "2990;450"
         MousePointer    =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblBankReverse 
         Height          =   255
         Left            =   45
         TabIndex        =   10
         Top             =   360
         Width           =   1695
         BackColor       =   16777215
         Caption         =   "Reverse Transaction"
         Size            =   "2990;450"
         MousePointer    =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame fraBankTransactions 
      Height          =   900
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   1500
         Y1              =   600
         Y2              =   600
      End
      Begin MSForms.Label lblBankTran 
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   7
         Top             =   600
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Bank Transfer"
         Size            =   "2566;450"
         MousePointer    =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   1500
         Y1              =   360
         Y2              =   360
      End
      Begin MSForms.Label lblBankTran 
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   6
         Top             =   360
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Bank Payment"
         Size            =   "2566;450"
         MousePointer    =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblBankTran 
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Top             =   120
         Width           =   1455
         BackColor       =   14737632
         Caption         =   "Bank Receipt"
         Size            =   "2566;450"
         MousePointer    =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtSetFocus 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame fraDemand 
      Height          =   660
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1785
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   1760
         Y1              =   360
         Y2              =   360
      End
      Begin MSForms.Label lblCopy 
         Height          =   255
         Left            =   45
         TabIndex        =   2
         Top             =   120
         Width           =   1695
         BackColor       =   14737632
         Caption         =   "Copy Transaction"
         Size            =   "2990;450"
         MousePointer    =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblReverse 
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Top             =   360
         Width           =   1695
         BackColor       =   16777215
         Caption         =   "Reverse Transaction"
         Size            =   "2990;450"
         MousePointer    =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Index           =   2
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   1695
      BackColor       =   14737632
      Caption         =   "<< Bank Transactions"
      Size            =   "2990;873"
      MousePointer    =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   1920
      Width           =   2415
      BackColor       =   14737632
      Caption         =   "<< Copy SI, SR - Demands form | Copy PI, PP - Purchase form"
      Size            =   "4260;873"
      MousePointer    =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
      BackColor       =   14737632
      Caption         =   "<< Copy Bank Trans, Bank Transactions"
      Size            =   "2990;873"
      MousePointer    =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmPopUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CALLING_FORM         As String
Public TRANSACITON_ID_LONG   As Long
Public TRANSACTION_ID_STR    As String
Public TRANS_TYPE            As String

Private Sub cmdClose_Click(index As Integer)
   Unload Me
End Sub

Private Sub cmdDP_Posting_Click()
   If optDPDtRange.Value Then
      If txtDPDtRangeFrom.text = "" Then
         ShowMsgInTaskBar "Please enter the FROM date", "Y", "N"
         txtDPDtRangeFrom.SetFocus
         Exit Sub
      End If
      If txtDPDtRangeTo.text = "" Then
         ShowMsgInTaskBar "Please enter the TO date", "Y", "N"
         txtDPDtRangeTo.SetFocus
         Exit Sub
      End If
   End If
   
   If optDPSlNoRange.Value Then
      If txtDPSlRangeFrom.text = "" Then
         ShowMsgInTaskBar "Please enter the demand number", "Y", "N"
         txtDPSlRangeFrom.SetFocus
         Exit Sub
      End If
      If txtDPSlRangeTo.text = "" Then
         ShowMsgInTaskBar "Please enter the demand number", "Y", "N"
         txtDPSlRangeTo.SetFocus
         Exit Sub
      End If
   End If
   cmdDP_Posting.Enabled = False
   frmDemands3.PostDemands
   cmdDP_Posting.Enabled = True
   Unload Me
End Sub

Private Sub cmdPI_Posting_Click()
   If optPIDtRange.Value Then
      If CDate(txtDtRangeFrom.text) > CDate(txtDtRangeTo.text) Then
         ShowMsgInTaskBar "From date should not be after the To date.", "Y", "N"
         txtDtRangeTo.SetFocus
         Exit Sub
      End If
   End If
   If optSlNoRange.Value Then
      If UCase(Left(txtPlRangeFrom.text, 2)) <> "PI" And UCase(Left(txtPlRangeFrom.text, 2)) <> "PC" Then
         ShowMsgInTaskBar "Please type the prefix of the rang, such as PIxx or PCxx.", "Y", "N"
         txtPlRangeFrom.SetFocus
         Exit Sub
      End If
      If UCase(Left(txtPlRangeTo.text, 2)) <> "PI" And UCase(Left(txtPlRangeTo.text, 2)) <> "PC" Then
         ShowMsgInTaskBar "Please type the prefix of the rang, such as PIxx or PCxx.", "Y", "N"
         txtPlRangeTo.SetFocus
         Exit Sub
      End If
      If UCase(Left(txtPlRangeFrom.text, 2)) <> UCase(Left(txtPlRangeTo.text, 2)) Then
         ShowMsgInTaskBar "Please enter the same the prefix in the both filed.", "Y", "N"
         txtPlRangeTo.SetFocus
         Exit Sub
      End If
      If StrDigitVal(txtPlRangeFrom.text) > StrDigitVal(txtPlRangeTo.text) Then
         ShowMsgInTaskBar "From number should not be bigger then To number.", "Y", "N"
         txtPlRangeTo.SetFocus
         Exit Sub
      End If
   End If
   If CALLING_FORM = "POSTPURCHASE" Then
        frmPurchaseExpense.PostInvoice
        frmPurchaseExpense.chkSelectAllDemands.Value = 0
   ElseIf CALLING_FORM = UCase("ManagementFee") Then
        frmManagementFees.PostInvoice
        frmManagementFees.chkSelectAllDemands.Value = 0
   End If
   Unload Me
End Sub

Private Sub cmdPO_Posting_Click()
   If optPODtRange.Value Then
      If CDate(txtDtRangeFromPO.text) > CDate(txtDtRangeToPO.text) Then
         ShowMsgInTaskBar "From date should not be after the To date.", "Y", "N"
         txtDtRangeToPO.SetFocus
         Exit Sub
      End If
   End If
   If optPONoRange.Value Then
      If UCase(Left(txtPORangeFrom.text, 2)) <> "PO" Then txtPORangeFrom.text = "PO" & txtPORangeFrom.text
      If UCase(Left(txtPORangeTo.text, 2)) <> "PO" Then txtPORangeTo.text = "PO" & txtPORangeTo.text
      If StrDigitVal(txtPORangeFrom.text) > StrDigitVal(txtPORangeTo.text) Then
         ShowMsgInTaskBar "From number should not be bigger then To number.", "Y", "N"
         txtPORangeTo.SetFocus
         Exit Sub
      End If
   End If
   If CALLING_FORM = UCase("ManagementFee") Then
        frmManagementFees.PostInvoice
        frmManagementFees.chkSelectAllDemands.Value = 0
   Else
        frmPO.PostInvoice
        frmPO.chkSelectAllDemands.Value = 0
   End If
   Unload Me
End Sub

'Private Sub Command1_Click()
'   Unload Me
'End Sub

Private Sub Form_Activate()
   If CALLING_FORM <> "POSTDEMANDS" And CALLING_FORM <> "POSTPURCHASE" And CALLING_FORM <> "POSTPO" Then _
      txtSetFocus.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Unload Me
End Sub

Public Sub CallingFrom(szForm As String)
   szForm = UCase(szForm)
   CALLING_FORM = szForm

   If Left(szForm, 6) = "DEMAND" Or _
         Left(szForm, 8) = "PURCHASE" Or _
         Left(szForm, 6) = "LESSEE" Or _
         szForm = "CB_SR" Or szForm = "CB_SA" Then
      fraDemand.Left = 0
      fraDemand.Top = -80
      Me.Height = 660
      Me.Width = 1860
      fraDemand.Visible = True
   End If
   If szForm = "BANK" Then
      fraBankTransactions.Left = 0
      fraBankTransactions.Top = -80
      Me.Height = 915
      Me.Width = 1620
      fraBankTransactions.Visible = True
   End If
   If szForm = "BANKCOPY" Or szForm = "CB_BANK" Then
      fraBankCopy.Left = 0
      fraBankCopy.Top = -80
      Me.Height = 660
      Me.Width = 1860
      fraBankCopy.Visible = True
   End If
   If szForm = "POSTDEMANDS" Then
      fraDemandPosting.Left = 0
      fraDemandPosting.Top = -80
      Me.Height = 3735
      Me.Width = 2610
      fraDemandPosting.Visible = True
   End If
   If szForm = "POSTPURCHASE" Then
      fraPurchasePosting.Left = 0
      fraPurchasePosting.Top = -80
      Me.Height = 3015
      Me.Width = 3105
      fraPurchasePosting.Visible = True
      optSelPO.Caption = "Selected Purchase Order"
   End If
   If szForm = "POSTPO" Then
      fraPurchaseOrder.Left = 0
      fraPurchaseOrder.Top = -80
      Me.Height = 3015
      Me.Width = 3105
      fraPurchaseOrder.Visible = True
      optSelPO.Caption = "Selected Purchase Order"
   End If
   If szForm = UCase("ManagementFee") Then
      fraPurchaseOrder.Left = 0
      fraPurchaseOrder.Top = -80
      Me.Height = 3015
      Me.Width = 3105
      fraPurchaseOrder.Visible = True
      optSelPO.Caption = "Selected Management Fee"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   fraDemand.Visible = False
   fraBankTransactions.Visible = False
End Sub

Private Sub lblBankCopy_Click()
   If CALLING_FORM = "CB_BANK" Then
      frmCashbook.CopyBankTran
      frmCashbook.cmdCopy.SetFocus
   Else
      frmBankTransactions.CopyTransaction
      frmBankTransactions.cmdCopy.SetFocus
   End If
End Sub

Private Sub lblBankCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblBankCopy.BackColor = &HE0E0E0
   lblBankReverse.BackColor = vbWhite
End Sub

Private Sub lblBankReverse_Click()
   If CALLING_FORM = "CB_BANK" Then
      frmCashbook.CopyRevTransaction
      frmCashbook.cmdCopy.SetFocus
   Else
      frmBankTransactions.CopyRevTransaction
      frmBankTransactions.cmdCopy.SetFocus
   End If
End Sub

Private Sub lblBankReverse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblBankCopy.BackColor = vbWhite
   lblBankReverse.BackColor = &HE0E0E0
End Sub

Private Sub lblBankTran_Click(index As Integer)
   If index = 0 Or index = 1 Then
      Load frmBankTranEdit
      frmBankTranEdit.FrmBankTranEdit_CALLING_FROM = frmBankTransactions.Name

      frmBankTranEdit.Caption = "Add " & lblBankTran(index).Caption

      frmBankTranEdit.Show
   End If
   If index = 2 Then
      Load frmBankTransfer
      frmBankTransfer.FrmBankTransfer_CALLING_FROM = "frmBankTransactions"

      frmBankTransfer.Show
   End If

   frmBankTransactions.Enabled = False
End Sub

Private Sub lblBankTran_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Integer

   lblBankTran(index).BackColor = &HE0E0E0
   For i = 0 To 2
      If i <> index Then
         lblBankTran(i).BackColor = vbWhite
      End If
   Next i
End Sub

Private Sub lblReverse_Click()
   Dim lTranID    As Long
   Dim szTranID   As String
   Dim szType     As String

   If CALLING_FORM = "DEMANDS" Or CALLING_FORM = "LESSEE_SI" Then
      lTranID = TRANSACITON_ID_LONG 'frmDemands3.SelectedID(szType)
      If TRANS_TYPE = "INV" Then
         Create_SC (lTranID)
      End If
      If TRANS_TYPE = "CRN" Then
         Create_SI (lTranID)
      End If
   End If
   If CALLING_FORM = "PURCHASE" Then
      szTranID = frmPurchaseExpense.flxPurchase.TextMatrix(frmPurchaseExpense.flxPurchase.row, 0)
      szType = Left(frmPurchaseExpense.flxPurchase.TextMatrix(frmPurchaseExpense.flxPurchase.row, 2), 2)
      If szType = "PI" Then
         Create_PC szTranID
      End If
      If szType = "PC" Then
         Create_PI szTranID
      End If
   End If
   If CALLING_FORM = "DEMANDRECEIPT" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "DEMAND_RECEIPT_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "LESSEE_SR" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "LESSEE_RECEIPT_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "CB_SR" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "CB_SR_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "DEMAND_SRR" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "DEMAND_SRR_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "LESSEE_SRR" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "LESSEE_SRR_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "DEMAND_SA" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "DEMAND_SA_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "LESSEE_SA" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "LESSEE_SA_REVERSE"
      frmCopyTransaction.Show 1
   End If

   If CALLING_FORM = "PURCHASE_PAYMENT" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "PURCHASE_PAYMENT_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "PURCHASE_PAYMENT_PPR" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "PAYMENT_REFUND_REVERSE"
      frmCopyTransaction.Show 1
   End If
   If CALLING_FORM = "PURCHASE_PAYMENT_ACCOUNT" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = "PAYMENT_ACCOUNT_REVERSE"
      frmCopyTransaction.Show 1
   End If
End Sub

Private Sub lblCopy_Click()
   Dim lTranID    As Long
   Dim szTranID   As String

   If CALLING_FORM = "DEMANDS" Or CALLING_FORM = "LESSEE_SI" Then
      lTranID = TRANSACITON_ID_LONG 'frmDemands3.SelectedID(szType)
      If TRANS_TYPE = "INV" Then
         Create_SI lTranID
      End If
      If TRANS_TYPE = "CRN" Then
         Create_SC lTranID
      End If
   End If

   If CALLING_FORM = "PURCHASE" Then
      szTranID = frmPurchaseExpense.flxPurchase.TextMatrix(frmPurchaseExpense.flxPurchase.row, 0)
      TRANS_TYPE = Left(frmPurchaseExpense.flxPurchase.TextMatrix(frmPurchaseExpense.flxPurchase.row, 2), 2)
      If TRANS_TYPE = "PI" Then
         Create_PI szTranID
      End If
      If TRANS_TYPE = "PC" Then
         Create_PC szTranID
      End If
   End If

   If CALLING_FORM = "DEMANDRECEIPT" Or _
         Right(CALLING_FORM, 3) = "_SR" Or _
         Right(CALLING_FORM, 4) = "_SRR" Or _
         Right(CALLING_FORM, 3) = "_SA" Or _
         CALLING_FORM = "PURCHASE_PAYMENT" Or _
         CALLING_FORM = "PURCHASE_PAYMENT_PPR" Or _
         CALLING_FORM = "PURCHASE_PAYMENT_ACCOUNT" Then
      Load frmCopyTransaction
      frmCopyTransaction.CALLING_FORM = CALLING_FORM
      frmCopyTransaction.Show 1
   End If
End Sub

Private Sub lblCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblCopy.BackColor = &HE0E0E0
   lblReverse.BackColor = vbWhite
End Sub

Private Sub lblReverse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblCopy.BackColor = vbWhite
   lblReverse.BackColor = &HE0E0E0
End Sub

Private Sub optDPDtRange_Click()
   If optDPDtRange.Value Then
      fraDPDtRange.Enabled = True
      If txtDPDtRangeFrom.text = "" Then txtDPDtRangeFrom.text = "01/01/2000"
      If txtDPDtRangeTo.text = "" Then txtDPDtRangeTo.text = Format(Now, "dd/mm/yyyy")
      txtDPDtRangeFrom.SetFocus
   End If
End Sub

Private Sub optDPSlNoRange_Click()
   If optDPSlNoRange.Value Then
      fraDPSlRange.Enabled = True
      txtDPSlRangeFrom.SetFocus
      If Not optSC.Value Then optSI.Value = True
   Else
      optSI.Value = False
      optSC.Value = False
   End If
End Sub

Private Sub optPIDtRange_Click()
   If optPIDtRange.Value Then
      fraPIDtRange.Enabled = True
      If txtDtRangeFrom.text = "" Then txtDtRangeFrom.text = "01/01/2000"
      If txtDtRangeTo.text = "" Then txtDtRangeTo.text = Format(Now, "dd/mm/yyyy")
      txtDtRangeFrom.SetFocus
   End If
End Sub

Private Sub optPODtRange_Click()
   'issue  469
   'added by anol 07 Jan 2015
   'test box was not enabled
   If optPODtRange.Value Then
      txtDtRangeFromPO.Enabled = True
      txtDtRangeToPO.Enabled = True
      Frame2.Enabled = True
   Else
      txtDtRangeFromPO.Enabled = False
      txtDtRangeToPO.Enabled = False
      Frame2.Enabled = False
   End If
   
End Sub

Private Sub optSlNoRange_Click()
   If optSlNoRange.Value Then
      fraPlRange.Enabled = True
      txtPlRangeFrom.SetFocus
   End If
End Sub

Private Sub txtDPDtRangeFrom_Change()
   TextBoxChangeDate txtDPDtRangeFrom
End Sub

Private Sub txtDPDtRangeFrom_GotFocus()
   SelTxtInCtrl txtDPDtRangeFrom
End Sub

Private Sub txtDPDtRangeFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDPDtRangeFrom, KeyAscii
End Sub

Private Sub txtDPDtRangeFrom_LostFocus()
   TextBoxFormatDate txtDPDtRangeFrom
End Sub

Private Sub txtDPDtRangeTo_Change()
   TextBoxChangeDate txtDPDtRangeTo
End Sub

Private Sub txtDPDtRangeTo_GotFocus()
   SelTxtInCtrl txtDPDtRangeTo
End Sub

Private Sub txtDPDtRangeTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDPDtRangeTo, KeyAscii
End Sub

Private Sub txtDPDtRangeTo_LostFocus()
   TextBoxFormatDate txtDPDtRangeTo
End Sub

Private Sub txtDPSlRangeFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtDPSlRangeTo.SetFocus
    End If
   DigitTextKeyPress txtDPSlRangeFrom, KeyAscii, 0
End Sub

Private Sub txtDPSlRangeTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDP_Posting.SetFocus
    End If
   DigitTextKeyPress txtDPSlRangeTo, KeyAscii, 0
End Sub

Private Sub txtDtRangeFrom_Change()
   TextBoxChangeDate txtDtRangeFrom
End Sub

Private Sub txtDtRangeFrom_GotFocus()
   SelTxtInCtrl txtDtRangeFrom
End Sub

Private Sub txtDtRangeFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDtRangeFrom, KeyAscii
End Sub

Private Sub txtDtRangeFrom_LostFocus()
   TextBoxFormatDate txtDtRangeFrom
End Sub

Private Sub txtDtRangeFromPO_Change()
   'anol 08 Jan 2015
   TextBoxChangeDate txtDtRangeFromPO
End Sub

Private Sub txtDtRangeFromPO_GotFocus()
   'anol 08 Jan 2015
   SelTxtInCtrl txtDtRangeFromPO
End Sub

Private Sub txtDtRangeFromPO_KeyPress(KeyAscii As Integer)
   'anol 08 Jan 2015
   TextBoxKeyPrsDate txtDtRangeFromPO, KeyAscii
End Sub

Private Sub txtDtRangeFromPO_LostFocus()
   'anol 08 Jan 2015
   TextBoxFormatDate txtDtRangeFromPO
End Sub

Private Sub txtDtRangeTo_Change()
   TextBoxChangeDate txtDtRangeTo
End Sub

Private Sub txtDtRangeTo_GotFocus()
   SelTxtInCtrl txtDtRangeTo
End Sub

Private Sub txtDtRangeTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDtRangeTo, KeyAscii
End Sub

Private Sub txtDtRangeTo_LostFocus()
   cmdPI_Posting.SetFocus
   TextBoxFormatDate txtDtRangeTo
End Sub

Private Sub txtDtRangeToPO_Change()
      'Anol 08 Jan 2015
      TextBoxChangeDate txtDtRangeToPO
End Sub

Private Sub txtDtRangeToPO_GotFocus()
  'Anol 08 Jan 2015
      SelTxtInCtrl txtDtRangeToPO
End Sub

Private Sub txtDtRangeToPO_KeyPress(KeyAscii As Integer)
  'Anol 08 Jan 2015
    TextBoxKeyPrsDate txtDtRangeToPO, KeyAscii
End Sub

Private Sub txtDtRangeToPO_LostFocus()
  'Anol 08 Jan 2015
       TextBoxFormatDate txtDtRangeToPO
End Sub

Private Sub txtPlRangeFrom_GotFocus()
   SelTxtInCtrl txtPlRangeFrom
End Sub

Private Sub txtPlRangeFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPlRangeTo.SetFocus
    End If
End Sub

Private Sub txtPlRangeTo_GotFocus()
   SelTxtInCtrl txtPlRangeTo
End Sub

Private Sub txtPlRangeTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPI_Posting.SetFocus
    End If
End Sub

Private Sub txtPlRangeTo_LostFocus()
   cmdPI_Posting.SetFocus
End Sub

Private Sub txtPORangeFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPORangeTo.SetFocus
    End If
End Sub

Private Sub txtPORangeTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPO_Posting.SetFocus
    End If
End Sub

Private Sub txtSetFocus_LostFocus()
   If CALLING_FORM <> "POSTDEMANDS" And CALLING_FORM <> "POSTPURCHASE" And CALLING_FORM <> "POSTPO" And CALLING_FORM <> "MANAGEMENTFEE" Then
      Unload Me
   End If
End Sub

Private Sub Create_SC(lSI As Long)
   Dim adoConn             As New ADODB.Connection
   Dim adoRstSplitDemand   As New ADODB.Recordset
   Dim adoRstDemandRec     As New ADODB.Recordset
   Dim adoRstNewDmd        As New ADODB.Recordset
   Dim adoRstNewSpDmd      As New ADODB.Recordset
   Dim szSQL               As String
   Dim lNxtDmdID           As Long
   Dim iChildId            As Integer

   On Error GoTo CatchError

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM DemandRecords WHERE DEMANDID = " & lSI & ""
   adoRstDemandRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   szSQL = "SELECT * FROM DemandSplitRecords WHERE DEMANDID = " & lSI & ""
   adoRstSplitDemand.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   lNxtDmdID = NextRef(adoConn, "DEMAND_REF")
   szSQL = "SELECT * FROM DemandRecords;"

   With adoRstNewDmd
      .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
      .AddNew
      .Fields.Item("DemandID").Value = lNxtDmdID
      .Fields.Item("SageAccountNumber").Value = adoRstDemandRec.Fields.Item("SageAccountNumber").Value
      .Fields.Item("TenantCompanyName").Value = adoRstDemandRec.Fields.Item("TenantCompanyName").Value
      .Fields.Item("UnitNumber").Value = adoRstDemandRec.Fields.Item("UnitNumber").Value
      .Fields.Item("Source").Value = 1
      .Fields.Item("TransactionType").Value = 2
      .Fields.Item("IssueDate").Value = Format(Now, "dd/mm/yyyy")
      .Fields.Item("SageText").Value = adoRstDemandRec.Fields.Item("DemandID").Value
      .Fields.Item("SageText").Value = "S/L " & adoRstDemandRec.Fields.Item("SageAccountNumber").Value
      .Fields.Item("IsPrinted").Value = False
      .Fields.Item("Spare1").Value = adoRstDemandRec.Fields.Item("Spare1").Value
      .Fields.Item("LeaseRef").Value = adoRstDemandRec.Fields.Item("LeaseRef").Value
      .Fields.Item("AdjTag").Value = adoRstDemandRec.Fields.Item("AdjTag").Value
      .Fields.Item("Details").Value = adoRstDemandRec.Fields.Item("Details").Value
      .Fields.Item("DmdSlNo").Value = SlNumber("SC", "DemandRecords", adoConn)
      .Update
      .Close
   End With
   adoRstDemandRec.Close

   szSQL = "SELECT * FROM DemandSplitRecords;"

   With adoRstNewSpDmd
      .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
      iChildId = 0

      While Not adoRstSplitDemand.EOF
         .AddNew
         .Fields.Item("DSR").Value = UniqueID()
         iChildId = iChildId + 1
         .Fields.Item("SplitID").Value = iChildId
         .Fields.Item("DemandId").Value = lNxtDmdID
         .Fields.Item("A_M").Value = "C"
         .Fields.Item("NominalCodeforAmount").Value = adoRstSplitDemand.Fields.Item("NominalCodeforAmount").Value
         .Fields.Item("NominalNameforAmount").Value = adoRstSplitDemand.Fields.Item("NominalNameforAmount").Value
         .Fields.Item("NominalCodeForVAT").Value = adoRstSplitDemand.Fields.Item("NominalCodeForVAT").Value
         .Fields.Item("NominalNameforVAT").Value = adoRstSplitDemand.Fields.Item("NominalNameforVAT").Value
         .Fields.Item("NominalCodeForTotal").Value = adoRstSplitDemand.Fields.Item("NominalCodeForTotal").Value
         .Fields.Item("NominalNameforTotal").Value = adoRstSplitDemand.Fields.Item("NominalNameforTotal").Value
         .Fields.Item("Amount").Value = adoRstSplitDemand.Fields.Item("Amount").Value
         .Fields.Item("VATAmount").Value = adoRstSplitDemand.Fields.Item("VATAmount").Value
         .Fields.Item("TotalAmount").Value = adoRstSplitDemand.Fields.Item("TotalAmount").Value
         .Fields.Item("SageRef").Value = adoRstSplitDemand.Fields.Item("DSR").Value
         .Fields.Item("DueDate").Value = adoRstSplitDemand.Fields.Item("DueDate").Value
         .Fields.Item("VATMonth").Value = adoRstSplitDemand.Fields.Item("VATMonth").Value
         .Fields.Item("TypeOfDemand").Value = adoRstSplitDemand.Fields.Item("VATMonth").Value
         .Fields.Item("description").Value = adoRstSplitDemand.Fields.Item("description").Value
         .Fields.Item("DemandStatement").Value = adoRstSplitDemand.Fields.Item("DemandStatement").Value
         .Fields.Item("VAT_CODE").Value = adoRstSplitDemand.Fields.Item("VAT_CODE").Value
         .Fields.Item("DateFrom").Value = adoRstSplitDemand.Fields.Item("DateFrom").Value
         .Fields.Item("DateTo").Value = adoRstSplitDemand.Fields.Item("DateTo").Value
         .Fields.Item("SageDepartment").Value = adoRstSplitDemand.Fields.Item("SageDepartment").Value
         .Fields.Item("FrequencyID").Value = adoRstSplitDemand.Fields.Item("FrequencyID").Value
         .Fields.Item("ChargingFigure").Value = adoRstSplitDemand.Fields.Item("ChargingFigure").Value
         .Fields.Item("AdiComment").Value = adoRstSplitDemand.Fields.Item("AdiComment").Value
         .Fields.Item("ChargingMethod").Value = adoRstSplitDemand.Fields.Item("ChargingMethod").Value
         .Fields.Item("JobID").Value = adoRstSplitDemand.Fields.Item("JobID").Value
         .Update

         adoRstSplitDemand.MoveNext
      Wend
   End With

   Set adoRstDemandRec = Nothing
   Set adoRstSplitDemand = Nothing
   Set adoRstNewDmd = Nothing
   Set adoRstNewSpDmd = Nothing

   If CALLING_FORM = "DEMANDS" Then
      frmDemands3.LoadFlxGrid frmDemands3.flxDemands, False, adoConn, ""
      frmDemands3.flxDemands.row = 0
      frmDemands3.flxDemands.col = 0
   End If

'  EXPORT all Invoices or Demands into tlbReceipt table *********************************************
   MigrateInvIntoReceipt adoConn

   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Transaction has been copied successfully.", "Y", "P"
   txtSetFocus_LostFocus
   Exit Sub

CatchError:
   ShowMsgInTaskBar "System could not copy the transaction.", "Y", "N"
   Set adoConn = Nothing
End Sub

Private Sub Create_SI(lSI As Long)
   Dim adoConn             As New ADODB.Connection
   Dim adoRstSplitDemand   As New ADODB.Recordset
   Dim adoRstDemandRec     As New ADODB.Recordset
   Dim adoRstNewDmd        As New ADODB.Recordset
   Dim adoRstNewSpDmd      As New ADODB.Recordset
   Dim szSQL               As String
   Dim lNxtDmdID           As Long
   Dim iChildId            As Integer

   On Error GoTo CatchError

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM DemandRecords WHERE DEMANDID = " & lSI & ""
   adoRstDemandRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   szSQL = "SELECT * FROM DemandSplitRecords WHERE DEMANDID = " & lSI & ""
   adoRstSplitDemand.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   lNxtDmdID = NextRef(adoConn, "DEMAND_REF")
   szSQL = "SELECT * FROM DemandRecords;"

   With adoRstNewDmd
      .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
      .AddNew
      .Fields.Item("DemandID").Value = lNxtDmdID
      .Fields.Item("SageAccountNumber").Value = adoRstDemandRec.Fields.Item("SageAccountNumber").Value
      .Fields.Item("TenantCompanyName").Value = adoRstDemandRec.Fields.Item("TenantCompanyName").Value
      .Fields.Item("UnitNumber").Value = adoRstDemandRec.Fields.Item("UnitNumber").Value
      .Fields.Item("Source").Value = 1
      .Fields.Item("TransactionType").Value = 1
      .Fields.Item("IssueDate").Value = Format(Now, "dd/mm/yyyy")
      .Fields.Item("SageText").Value = adoRstDemandRec.Fields.Item("DemandID").Value
      .Fields.Item("SageText").Value = "S/L " & adoRstDemandRec.Fields.Item("SageAccountNumber").Value
      .Fields.Item("IsPrinted").Value = False
      .Fields.Item("Spare1").Value = adoRstDemandRec.Fields.Item("Spare1").Value
      .Fields.Item("LeaseRef").Value = adoRstDemandRec.Fields.Item("LeaseRef").Value
      .Fields.Item("AdjTag").Value = adoRstDemandRec.Fields.Item("AdjTag").Value
      .Fields.Item("Details").Value = adoRstDemandRec.Fields.Item("Details").Value
      .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoConn)
      .Update
      .Close
   End With
   adoRstDemandRec.Close

   szSQL = "SELECT * FROM DemandSplitRecords;"

   With adoRstNewSpDmd
      .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
      iChildId = 0

      While Not adoRstSplitDemand.EOF
         .AddNew
         .Fields.Item("DSR").Value = UniqueID()
         iChildId = iChildId + 1
         .Fields.Item("SplitID").Value = iChildId
         .Fields.Item("DemandId").Value = lNxtDmdID
         .Fields.Item("A_M").Value = "C"
         .Fields.Item("NominalCodeforAmount").Value = adoRstSplitDemand.Fields.Item("NominalCodeforAmount").Value
         .Fields.Item("NominalNameforAmount").Value = adoRstSplitDemand.Fields.Item("NominalNameforAmount").Value
         .Fields.Item("NominalCodeForVAT").Value = adoRstSplitDemand.Fields.Item("NominalCodeForVAT").Value
         .Fields.Item("NominalNameforVAT").Value = adoRstSplitDemand.Fields.Item("NominalNameforVAT").Value
         .Fields.Item("NominalCodeForTotal").Value = adoRstSplitDemand.Fields.Item("NominalCodeForTotal").Value
         .Fields.Item("NominalNameforTotal").Value = adoRstSplitDemand.Fields.Item("NominalNameforTotal").Value
         .Fields.Item("Amount").Value = adoRstSplitDemand.Fields.Item("Amount").Value
         .Fields.Item("VATAmount").Value = adoRstSplitDemand.Fields.Item("VATAmount").Value
         .Fields.Item("TotalAmount").Value = adoRstSplitDemand.Fields.Item("TotalAmount").Value
         .Fields.Item("SageRef").Value = adoRstSplitDemand.Fields.Item("DSR").Value
         .Fields.Item("DueDate").Value = adoRstSplitDemand.Fields.Item("DueDate").Value
         .Fields.Item("VATMonth").Value = adoRstSplitDemand.Fields.Item("VATMonth").Value
         .Fields.Item("TypeOfDemand").Value = adoRstSplitDemand.Fields.Item("TypeOfDemand").Value
         .Fields.Item("description").Value = adoRstSplitDemand.Fields.Item("description").Value
         .Fields.Item("DemandStatement").Value = adoRstSplitDemand.Fields.Item("DemandStatement").Value
         .Fields.Item("VAT_CODE").Value = adoRstSplitDemand.Fields.Item("VAT_CODE").Value
         .Fields.Item("DateFrom").Value = adoRstSplitDemand.Fields.Item("DateFrom").Value
         .Fields.Item("DateTo").Value = adoRstSplitDemand.Fields.Item("DateTo").Value
         .Fields.Item("SageDepartment").Value = adoRstSplitDemand.Fields.Item("SageDepartment").Value
         .Fields.Item("FrequencyID").Value = adoRstSplitDemand.Fields.Item("FrequencyID").Value
         .Fields.Item("ChargingFigure").Value = adoRstSplitDemand.Fields.Item("ChargingFigure").Value
         .Fields.Item("AdiComment").Value = adoRstSplitDemand.Fields.Item("AdiComment").Value
         .Fields.Item("ChargingMethod").Value = adoRstSplitDemand.Fields.Item("ChargingMethod").Value
         .Fields.Item("JobID").Value = adoRstSplitDemand.Fields.Item("JobID").Value
         .Update

         adoRstSplitDemand.MoveNext
      Wend
   End With

   Set adoRstDemandRec = Nothing
   Set adoRstSplitDemand = Nothing
   Set adoRstNewDmd = Nothing
   Set adoRstNewSpDmd = Nothing

   If CALLING_FORM = "DEMANDS" Then
      frmDemands3.LoadFlxGrid frmDemands3.flxDemands, False, adoConn, ""
      frmDemands3.flxDemands.row = 0
      frmDemands3.flxDemands.col = 0
   End If

'  EXPORT all Invoices or Demands into tlbReceipt table *********************************************
   MigrateInvIntoReceipt adoConn

   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Transaction has been copied successfully.", "Y", "P"
   txtSetFocus_LostFocus
   Exit Sub

CatchError:
   ShowMsgInTaskBar "System could not copy the transaction.", "Y", "N"
   Set adoConn = Nothing
End Sub

Private Sub Create_PI(szTranID As String)
   Dim adoConn       As New ADODB.Connection
   Dim adoPIHeader   As New ADODB.Recordset
   Dim adoPISplit    As New ADODB.Recordset
   Dim adoSrcPIHd    As New ADODB.Recordset
   Dim adoSrcPISp    As New ADODB.Recordset
   Dim szSQL      As String
   Dim iRow       As Integer
   Dim uID        As String
   Dim lT_ID      As Long
   Dim lSlNumber  As Long

   adoConn.Open getConnectionString

'  ***************************************************************************************************
'           SAVING HEADER PART OF THE PURCHASE INVOICE                                               '
'  ***************************************************************************************************
   szSQL = "SELECT * " & _
           "FROM   tblPurInv " & _
           "WHERE  MY_ID = '" & szTranID & "';"
   adoSrcPIHd.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   szSQL = "SELECT * FROM tblPurInvSRec WHERE ParentID = '" & szTranID & "';"
   adoSrcPISp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT * FROM tblPurInv;"
   adoPIHeader.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   With adoPIHeader
      .AddNew
      uID = UniqueID()
      .Fields.Item("MY_ID").Value = uID

      lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
      .Fields.Item("SlNumber").Value = lSlNumber
      .Fields.Item("SUPP_AC").Value = adoSrcPIHd.Fields.Item("SUPP_AC").Value
      .Fields.Item("TRAN_DATE").Value = Format(Now, "DD/MMMM/YYYY")
      .Fields.Item("TransactionType").Value = 6
      .Fields.Item("INV_NO").Value = "C - " & adoSrcPIHd.Fields.Item("INV_NO").Value
      .Fields.Item("TOTAL_AMOUNT").Value = adoSrcPIHd.Fields.Item("TOTAL_AMOUNT").Value
      .Fields.Item("TTP").Value = CByte(TransactionTakePlace("TTP", "PURCHASE INVOICE", adoConn))
      .Fields.Item("History").Value = False
      .Fields.Item("TrfPayment").Value = True
      .Fields.Item("PropertyID").Value = adoSrcPIHd.Fields.Item("PropertyID").Value
      .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")

      .Update
      .Close
   End With

'  ***************************************************************************************************
'           B4 SAVING SPLITS, THE PI IS EXPORTED TO PAYMENT TABLE                                    '
'  ***************************************************************************************************
   
   szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
   adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
   adoPIHeader.Close

   szSQL = "SELECT * FROM tlbPayment"
   'If cmdEdit(1).Enabled = false Then edit mode

'Debug.Print szSQL
   adoPIHeader.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   With adoPIHeader
      .AddNew
      !transactionID = lT_ID
      !Pi = uID
      !Type = 6
      !SageAccountNumber = adoSrcPIHd.Fields.Item("SUPP_AC").Value
      !PDate = Format(Now, "DD MMMM YYYY")
      !dDate = Format(Now, "DD MMMM YYYY")
      !ref = "C - " & adoSrcPIHd.Fields.Item("INV_NO").Value
      !ExtRef = !ref
      !amount = CCur(adoSrcPIHd.Fields.Item("TOTAL_AMOUNT").Value)
      !OSAmount = !amount
      !PaymentView = True
      !Details = adoSrcPISp.Fields.Item("DESCRIPTION").Value
      !unitid = adoSrcPIHd.Fields.Item("PropertyID").Value
      !SlNumber = lSlNumber
      !AdjTag = "N"

      .Update
      .Close
   End With
   adoSrcPIHd.Close

'  ***************************************************************************************************
'           SAVING SPLITS OF THE PURCHASE INVOICE
'  ***************************************************************************************************
   szSQL = "SELECT * FROM tblPurInvSRec"
   adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

'Add New Records
   While Not adoSrcPISp.EOF
      With adoPISplit
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ParentID").Value = uID
         .Fields.Item("TRAN_ID").Value = adoSrcPISp.Fields.Item("TRAN_ID").Value
         .Fields.Item("TRANS").Value = adoSrcPISp.Fields.Item("TRANS").Value
         .Fields.Item("UNIT_ID").Value = adoSrcPISp.Fields.Item("UNIT_ID").Value
         .Fields.Item("NOMINAL_CODE").Value = adoSrcPISp.Fields.Item("NOMINAL_CODE").Value
         .Fields.Item("DEPT_ID").Value = adoSrcPISp.Fields.Item("DEPT_ID").Value
         .Fields.Item("JOB_ID").Value = adoSrcPISp.Fields.Item("JOB_ID").Value
         .Fields.Item("COST_CODE").Value = adoSrcPISp.Fields.Item("COST_CODE").Value
         .Fields.Item("description").Value = adoSrcPISp.Fields.Item("description").Value
         .Fields.Item("NET_AMOUNT").Value = CCur(adoSrcPISp.Fields.Item("NET_AMOUNT").Value)
         .Fields.Item("TAX_CODE").Value = adoSrcPISp.Fields.Item("TAX_CODE").Value
         .Fields.Item("VAT").Value = CCur(adoSrcPISp.Fields.Item("VAT").Value)
         .Fields.Item("ScheduleID").Value = adoSrcPISp.Fields.Item("ScheduleID").Value
         .Fields.Item("TOTAL_AMOUNT").Value = CCur(adoSrcPISp.Fields.Item("TOTAL_AMOUNT").Value)

         .Update
      End With
      adoSrcPISp.MoveNext
   Wend
   adoPISplit.Close
   adoSrcPISp.Close

   ShowMsgInTaskBar "Transaction has been copied successfully.", "Y", "P"
   frmPurchaseExpense.LoadFlxPurchase adoConn
   frmPurchaseExpense.flxPurchase.SetFocus

   adoConn.Close

   Set adoSrcPIHd = Nothing
   Set adoSrcPISp = Nothing
   Set adoPISplit = Nothing
   Set adoPIHeader = Nothing
   Set adoConn = Nothing
End Sub

Private Sub Create_PC(szTranID As String)
   Dim adoConn       As New ADODB.Connection
   Dim adoPIHeader   As New ADODB.Recordset
   Dim adoPISplit    As New ADODB.Recordset
   Dim adoSrcPIHd    As New ADODB.Recordset
   Dim adoSrcPISp    As New ADODB.Recordset
   Dim szSQL      As String
   Dim iRow       As Integer
   Dim uID        As String
   Dim lT_ID      As Long
   Dim lSlNumber  As Long

   adoConn.Open getConnectionString

'  ***************************************************************************************************
'           SAVING HEADER PART OF THE PURCHASE INVOICE                                               '
'  ***************************************************************************************************
   szSQL = "SELECT * " & _
           "FROM   tblPurInv " & _
           "WHERE  MY_ID = '" & szTranID & "';"
   adoSrcPIHd.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT * FROM tblPurInvSRec WHERE ParentID = '" & szTranID & "';"
   adoSrcPISp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT * FROM tblPurInv;"
   adoPIHeader.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   With adoPIHeader
      .AddNew
      uID = UniqueID()
      .Fields.Item("MY_ID").Value = uID

      lSlNumber = SlNumber("PC", "tblPurInv", adoConn)
      .Fields.Item("SlNumber").Value = lSlNumber
      .Fields.Item("SUPP_AC").Value = adoSrcPIHd.Fields.Item("SUPP_AC").Value
      .Fields.Item("TRAN_DATE").Value = Format(Now, "DD/MMMM/YYYY")
      .Fields.Item("TransactionType").Value = 7
      .Fields.Item("INV_NO").Value = "C - " & adoSrcPIHd.Fields.Item("INV_NO").Value
      .Fields.Item("TOTAL_AMOUNT").Value = adoSrcPIHd.Fields.Item("TOTAL_AMOUNT").Value
      .Fields.Item("TTP").Value = CByte(TransactionTakePlace("TTP", "PURCHASE INVOICE", adoConn))
      .Fields.Item("History").Value = False
      .Fields.Item("TrfPayment").Value = True
      .Fields.Item("PropertyID").Value = adoSrcPIHd.Fields.Item("PropertyID").Value
      .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")

      .Update
      .Close
   End With

'  ***************************************************************************************************
'           B4 SAVING SPLITS, THE PI IS EXPORTED TO PAYMENT TABLE                                    '
'  ***************************************************************************************************
   
   szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
   adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
   adoPIHeader.Close

   szSQL = "SELECT * FROM tlbPayment"

'Debug.Print szSQL
   adoPIHeader.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   With adoPIHeader
      .AddNew
      !transactionID = lT_ID
      !Pi = uID
      !Type = 7
      !SageAccountNumber = adoSrcPIHd.Fields.Item("SUPP_AC").Value
      !PDate = Format(Now, "DD MMMM YYYY")
      !dDate = Format(Now, "DD MMMM YYYY")
      !ref = "C - " & adoSrcPIHd.Fields.Item("INV_NO").Value
      !ExtRef = !ref
      !amount = CCur(adoSrcPIHd.Fields.Item("TOTAL_AMOUNT").Value)
      !OSAmount = !amount
      !PaymentView = True
      !Details = adoSrcPISp.Fields.Item("DESCRIPTION").Value
      !unitid = adoSrcPIHd.Fields.Item("PropertyID").Value
      !SlNumber = lSlNumber
      !AdjTag = "N"

      .Update
      .Close
   End With
   adoSrcPIHd.Close

'  ***************************************************************************************************
'           SAVING SPLITS OF THE PURCHASE INVOICE
'  ***************************************************************************************************
   szSQL = "SELECT * FROM tblPurInvSRec"
   adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

'Add New Records
   While Not adoSrcPISp.EOF
      With adoPISplit
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ParentID").Value = uID
         .Fields.Item("TRAN_ID").Value = adoSrcPISp.Fields.Item("TRAN_ID").Value
         .Fields.Item("TRANS").Value = adoSrcPISp.Fields.Item("TRANS").Value
         .Fields.Item("UNIT_ID").Value = adoSrcPISp.Fields.Item("UNIT_ID").Value
         .Fields.Item("NOMINAL_CODE").Value = adoSrcPISp.Fields.Item("NOMINAL_CODE").Value
         .Fields.Item("DEPT_ID").Value = adoSrcPISp.Fields.Item("DEPT_ID").Value
         .Fields.Item("JOB_ID").Value = adoSrcPISp.Fields.Item("JOB_ID").Value
         .Fields.Item("COST_CODE").Value = adoSrcPISp.Fields.Item("COST_CODE").Value
         .Fields.Item("description").Value = adoSrcPISp.Fields.Item("description").Value
         .Fields.Item("NET_AMOUNT").Value = CCur(adoSrcPISp.Fields.Item("NET_AMOUNT").Value)
         .Fields.Item("TAX_CODE").Value = adoSrcPISp.Fields.Item("TAX_CODE").Value
         .Fields.Item("VAT").Value = CCur(adoSrcPISp.Fields.Item("VAT").Value)
         .Fields.Item("ScheduleID").Value = adoSrcPISp.Fields.Item("ScheduleID").Value
         .Fields.Item("TOTAL_AMOUNT").Value = CCur(adoSrcPISp.Fields.Item("TOTAL_AMOUNT").Value)

         .Update
      End With
      adoSrcPISp.MoveNext
   Wend
   adoPISplit.Close
   adoSrcPISp.Close

   ShowMsgInTaskBar "Transaction has been copied successfully.", "Y", "P"
   frmPurchaseExpense.LoadFlxPurchase adoConn
   frmPurchaseExpense.flxPurchase.SetFocus

   adoConn.Close

   Set adoSrcPIHd = Nothing
   Set adoSrcPISp = Nothing
   Set adoPISplit = Nothing
   Set adoPIHeader = Nothing
   Set adoConn = Nothing
End Sub
