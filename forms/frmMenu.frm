VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "This form has not been used it. Example of using this module is found S:\Samrat\Desktop\sOMEcODEs\non mouse popup menu"
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   5175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuContactScam 
         Caption         =   "Contact Scam"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbout_Click()
MsgBox "This tutorial has been brought to you by Nod Programming Inc. For further assistance contact Nod's website or click on the Contact Scam Menu.", , "Popupmenu Tutorial - Nod Programming Inc"
End Sub

Private Sub mnuContactScam_Click()
MsgBox "To contact Scam please send email to ndulge-me@rocketmail.com or write to the Nod webmaster at nod_programer@hotmail.com", , "Popupmenu Tutorial - Nod Programming Inc"
End Sub

Private Sub mnuExit_Click()
Unload frmMain
Unload Me
End
End Sub
