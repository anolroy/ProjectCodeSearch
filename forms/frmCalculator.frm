VERSION 5.00
Begin VB.Form frmCalculator 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   3420
   ClientLeft      =   7935
   ClientTop       =   3270
   ClientWidth     =   3810
   ForeColor       =   &H00000000&
   Icon            =   "frmCalculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   3810
   Begin VB.TextBox txtKeyCatch 
      Height          =   495
      Left            =   3240
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0."
      Top             =   120
      Width           =   3555
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2820
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2820
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2820
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2820
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2820
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2280
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2280
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1740
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1740
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1740
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1740
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1740
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   675
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1155
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdblResult           As Double
Private mdblSavedNumber      As Double
Private mstrDot              As String
Private mstrOp               As String
Private mstrDisplay          As String
Private mblnDecEntered       As Boolean
Private mblnOpPending        As Boolean
Private mblnNewEquals        As Boolean
Private mblnEqualsPressed    As Boolean
Private mintCurrKeyIndex     As Integer
Private bNotConsider         As Boolean

Private Sub cmdCalc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then MsgBox "Bingo!"
End Sub

Private Sub cmdCalc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   bNotConsider = True
End Sub

Private Sub Form_Load()
   bNotConsider = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim intIndex    As Integer
   
   Select Case KeyCode
      Case vbKeyBack:             intIndex = 0
      Case vbKeyDelete:           intIndex = 1
      Case vbKeyEscape:           intIndex = 2
      Case vbKey0, vbKeyNumpad0:  intIndex = 18
      Case vbKey1, vbKeyNumpad1:  intIndex = 13
      Case vbKey2, vbKeyNumpad2:  intIndex = 14
      Case vbKey3, vbKeyNumpad3:  intIndex = 15
      Case vbKey4, vbKeyNumpad4:  intIndex = 8
      Case vbKey5, vbKeyNumpad5:  intIndex = 9
      Case vbKey6, vbKeyNumpad6:  intIndex = 10
      Case vbKey7, vbKeyNumpad7:  intIndex = 3
      Case vbKey8, vbKeyNumpad8:  intIndex = 4
      Case vbKey9, vbKeyNumpad9:  intIndex = 5
      Case vbKeyDecimal:          intIndex = 20
      Case vbKeyAdd:              intIndex = 21
      Case vbKeySubtract:         intIndex = 16
      Case vbKeyMultiply:         intIndex = 11
      Case vbKeyDivide:           intIndex = 6
      Case Else:                  Exit Sub
   End Select

   bNotConsider = True
   cmdCalc(intIndex).SetFocus

   cmdCalc_Click intIndex
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim intIndex    As Integer

   Select Case Chr$(KeyAscii)
      Case "S", "s":  intIndex = 7
      Case "P", "p":  intIndex = 12
      Case "R", "r":  intIndex = 17
      Case "X", "x":  intIndex = 11
      Case "=":       intIndex = 22
      Case Else:      Exit Sub
   End Select

   bNotConsider = True
   cmdCalc(intIndex).SetFocus

   cmdCalc_Click intIndex
End Sub

Private Sub cmdCalc_Click(Index As Integer)
   If Not bNotConsider Then Exit Sub

   Dim strPressedKey   As String

   mintCurrKeyIndex = Index

   If mstrDisplay = "ERROR" Then mstrDisplay = ""

   strPressedKey = cmdCalc(Index).Caption

   Select Case strPressedKey
      Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
         If mblnOpPending Then
            mstrDisplay = ""
            mblnOpPending = False
         End If
         If mblnEqualsPressed Then
            mstrDisplay = ""
            mblnEqualsPressed = False
         End If
         mstrDisplay = mstrDisplay & strPressedKey
      Case "."
         If mblnOpPending Then
            mstrDisplay = ""
            mblnOpPending = False
         End If
         If mblnEqualsPressed Then
            mstrDisplay = ""
            mblnEqualsPressed = False
         End If
         If InStr(mstrDisplay, ".") > 0 Then
            Beep
         Else
            mstrDisplay = mstrDisplay & strPressedKey
         End If
      Case "+", "-", "X", "/"
         mdblResult = Val(mstrDisplay)
         mstrOp = strPressedKey
         mblnOpPending = True
         mblnDecEntered = False
         mblnNewEquals = True
      Case "%"
         mdblSavedNumber = (Val(mstrDisplay) / 100) * mdblResult
         mstrDisplay = Format$(mdblSavedNumber)
      Case "="
         If mblnNewEquals Then
            mdblSavedNumber = Val(mstrDisplay)
            mblnNewEquals = False
         End If
         Select Case mstrOp
            Case "+"
               mdblResult = mdblResult + mdblSavedNumber
            Case "-"
               mdblResult = mdblResult - mdblSavedNumber
            Case "X"
               mdblResult = mdblResult * mdblSavedNumber
            Case "/"
               If mdblSavedNumber = 0 Then
                   mstrDisplay = "ERROR"
               Else
                   mdblResult = mdblResult / mdblSavedNumber
               End If
            Case Else
               mdblResult = Val(mstrDisplay)
         End Select
         If mstrDisplay <> "ERROR" Then
            mstrDisplay = Format$(mdblResult)
         End If
         mblnEqualsPressed = True
      Case "+/-"
         If mstrDisplay <> "" Then
            If Left$(mstrDisplay, 1) = "-" Then
               mstrDisplay = Right$(mstrDisplay, 2)
            Else
               mstrDisplay = "-" & mstrDisplay
            End If
         End If
      Case "Backspace"
         If Val(mstrDisplay) <> 0 Then
            mstrDisplay = Left$(mstrDisplay, Len(mstrDisplay) - 1)
            mdblResult = Val(mstrDisplay)
         End If
      Case "CE"
         mstrDisplay = ""
      Case "C"
         mstrDisplay = ""
         mdblResult = 0
         mdblSavedNumber = 0
      Case "1/x"
         If Val(mstrDisplay) = 0 Then
            mstrDisplay = "ERROR"
         Else
            mdblResult = Val(mstrDisplay)
            mdblResult = 1 / mdblResult
            mstrDisplay = Format$(mdblResult)
         End If
      Case "sqrt"
         If Val(mstrDisplay) < 0 Then
            mstrDisplay = "ERROR"
         Else
            mdblResult = Val(mstrDisplay)
            mdblResult = Sqr(mdblResult)
            mstrDisplay = Format$(mdblResult)
         End If
   End Select

   If mstrDisplay = "" Then
      txtDisplay = "0."
   Else
      mstrDot = IIf(InStr(mstrDisplay, ".") > 0, "", ".")
      txtDisplay = mstrDisplay & mstrDot
      If Left$(txtDisplay, 1) = "0" Then
         txtDisplay = Mid$(txtDisplay, 2)
      End If
   End If

   If txtDisplay = "." Then txtDisplay = "0."
   bNotConsider = False
   txtKeyCatch.SetFocus
End Sub

Private Sub txtKeyCatch_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      Form_KeyPress (61)
      txtKeyCatch.text = ""
   End If
End Sub
