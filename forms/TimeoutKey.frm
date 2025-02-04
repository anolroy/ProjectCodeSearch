VERSION 5.00
Begin VB.Form frmTimeoutKey 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "TimeoutKey.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Activate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "This program has timed out, please contact PCM Consulting Ltd for you reactivation key. Click OK to enter the key."
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmTimeoutKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheKey1

Public Sub Command1_Click()

Command1.Visible = False
Text2.Enabled = True
Command2.Visible = True
Command2.Enabled = True



'Either read the number of reactivations or use the year to verify
'For time being
Dim i As Integer
Dim i2 As Integer
Dim TimeoutNumber
Dim YearOfInstallation
YearOfInstallation = 2002

'Check installation year determine whether to do first write or rewrite, call appropriate function
Dim Year
Year = DatePart("yyyy", Date)
Dim Month
Month = DatePart("m", Date)
Dim Day
Day = DatePart("d", Date)
Dim Pos As Integer
Pos = 0

'Should assign appropriate timeout dates
If (Month > 11) Or ((Day >= 28) And Month = 11) Then
TimeoutNumber = Year - YearOfInstallation + 1
Else
TimeoutNumber = Year - YearOfInstallation
End If

i = 0
i2 = 0

Dim MaxLoop As Long

MaxLoop = 10

Dim TheNumber As Integer

TheNumber = 1

Dim a As Integer

a = 1

'Dim TheKey1 As Long

TheKey1 = 1


'key only 8 digets long
'TheKey1 = 50000000

'Dim TheKey2 As Long

'TheKey2 = 1


'key only 8 digets long
'TheKey2 = 50000000

Dim MEKArray(1000) As Integer

Dim CP As Integer
CP = 0

Dim BaseSkip As Integer
BaseSkip = 1

'Algorythm to generate key-must be duplicated
Dim TheFirstTenPrimes(10) As Integer

TheFirstTenPrimes(0) = 2
TheFirstTenPrimes(1) = 3
TheFirstTenPrimes(2) = 5
TheFirstTenPrimes(3) = 7
TheFirstTenPrimes(4) = 11
TheFirstTenPrimes(5) = 13
TheFirstTenPrimes(6) = 17
TheFirstTenPrimes(7) = 19
TheFirstTenPrimes(8) = 23
TheFirstTenPrimes(9) = 29

'Create MEKArray

For i = 0 To 1000

MEKArray(i) = TheFirstTenPrimes(i2)

'i = i + 1
i2 = i2 + 1

If i2 = 10 Then
i2 = 0
End If

Next i

'Text1.Text = MEKArray(999)
'TheNumber = 1 '81192435
'TheNumber = 2 '21146579
'TheNumber = 8
'TheNumber = 30
'TheNumber should equal the timeout number + a baseskip



TheNumber = TimeoutNumber

i = 1
i2 = 0

'Variables declared above
Do While (TheKey1 < 20000000)
TheKey1 = TheKey1 * MEKArray(i)

i = i + (TheNumber * i)
'For TheNumber = 30: i = 1, i = 31, i =
If i > 1000 Then
    i = 1 'needs to stay in 0 - 999
End If

Loop

'End of algorythm, Add 1 to the number of reactivations
'Number = Number + 1
'Write it back to encrypted file
Text2.text = TheKey1

End Sub

Public Sub Command2_Click()

    If Text2.text = TheKey1 Then
        Call RewriteTimeoutFile
        frmTimeoutKey.Hide
        Load frmlogin2
        frmlogin2.Show
        
    Else
        End
    End If

End Sub

