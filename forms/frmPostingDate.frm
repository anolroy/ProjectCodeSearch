VERSION 5.00
Begin VB.Form frmPostingDate 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   990
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPostingDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPostingDate 
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1395
      TabIndex        =   0
      Top             =   360
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date:"
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
      Index           =   42
      Left            =   195
      TabIndex        =   1
      Top             =   405
      Width           =   945
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   735
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3000
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   735
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmPostingDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public szCallingForm As String
Public szClientID    As String
Public szTransactionDate    As Date
Public szAllocationTransactionID    As String

Private Sub Form_Activate()
    If szCallingForm = "frmRevPayment" Or szCallingForm = "frmReverseAllocation" Then
        Label1(42).Caption = "Allocation Date:"
    Else
        Label1(42).Caption = "Posting Date:"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If szCallingForm = "frmDemands3" Then frmDemands3.Enabled = True
   If szCallingForm = "frmBatchDemands" Then frmBatchDemands.Enabled = True
   If szCallingForm = "frmPurchaseExpense" Then frmPurchaseExpense.Enabled = True
   If szCallingForm = "frmBankTranEdit" Then frmBankTranEdit.Enabled = True
   If szCallingForm = "frmBankTransfer" Then frmBankTransfer.Enabled = True
   If szCallingForm = "frmBRPreForm" Then frmBRPreForm.Enabled = True
   If szCallingForm = "frmBPPreForm" Then frmBPPreForm.Enabled = True
   If szCallingForm = "frmNJ_Entry" Then frmNJ_Entry.Enabled = True
   If szCallingForm = "frmPO2PI" Then frmPO2PI.Enabled = True
   'Resolved by BOSL
    'Below line added by anol 29 Mar 2015
    'issue 549: Demand receipts not working note 3
   If szCallingForm = "frmReceiptEdit" Then frmReceiptEdit.Enabled = True
   If szCallingForm = "frmBatchRpt" Then frmBatchRpt.Enabled = True
   If szCallingForm = "frmPaymentEdit" Then frmPaymentEdit.Enabled = True
   If szCallingForm = "frmRevPayment" Then frmRevPayment.Enabled = True
   If szCallingForm = "frmReverseAllocation" Then frmReverseAllocation.Enabled = True
End Sub







Private Sub txtPostingDate_Change()
   TextBoxChangeDate txtPostingDate
End Sub

Private Sub txtPostingDate_GotFocus()
   SelTxtInCtrl txtPostingDate
End Sub

Private Sub txtPostingDate_KeyPress(KeyAscii As Integer)
'added by anol because txtPostingDate.SetFocus was causing a problem
'Date 02 Dec 2014
   On Error Resume Next
   'end of addition
   If KeyAscii = 27 Then Unload Me                                'Escape

   Dim iRet As Integer

   If KeyAscii = 13 Then                                 'Enter
      'Resolved by BOSL
      'issue 468 : Posting Date
      'Modified by Anol 01 Sep 2014
      '   If txtPostingDate.text = "" Then Unload Me
        If IsDate(txtPostingDate.text) = False Then
            txtPostingDate.SetFocus
            If szCallingForm = "frmRevPayment" Or szCallingForm = "frmReverseAllocation" Then
                MsgBox "Please Enter a Valid Allocation Date."
            Else
                MsgBox "Please Enter a Valid Posting Date."
            End If
            Exit Sub
        End If

        Dim adoconn As New ADODB.Connection
        If szCallingForm = "frmRevPayment" Or szCallingForm = "frmReverseAllocation" Then
            iRet = DateDiff("d", szTransactionDate, txtPostingDate.text)
            If iRet < 0 Then
                MsgBox "Allocation date entered cannot be before the transaction date."
                Exit Sub
            Else
                'Now update the allocation transation with new transation date
                If MsgBox("Do you want to change the allocation date for this transaction?", vbYesNo, "Please confirm.") = vbYes Then
                    adoconn.Open getConnectionString
                    If szCallingForm = "frmRevPayment" Then
                        adoconn.Execute "Update Paytransactions set allocdate=#" & Format(txtPostingDate.text, "dd MMM yyyy") & "# where transactionID=" & szAllocationTransactionID & ""
                    ElseIf szCallingForm = "frmReverseAllocation" Then
                        adoconn.Execute "Update RptTransactions set allocdate=#" & Format(txtPostingDate.text, "dd MMM yyyy") & "# where transactionID=" & szAllocationTransactionID & ""
                    End If
                    adoconn.Close
                    Set adoconn = Nothing
                    MsgBox "Allocation date has been successfully amended for this transaction.", vbInformation, "Saved"
                    txtPostingDate.text = ""
                    Unload Me
                End If
            End If
        Else
              adoconn.Open getConnectionString
              iRet = IsPeriodStatus(txtPostingDate.text, szClientID, adoconn)         'Return value 0, 1, 9. 0 --> Close, 1 --> Open, 9 --> Not found
              adoconn.Close
              Set adoconn = Nothing
        
              If iRet = 9 Then
                 ShowMsgInTaskBar "The posting date entered does not fall in any period", "Y", "N"
                 txtPostingDate.SetFocus
                 Exit Sub
              End If
              If iRet = 0 Then
                 ShowMsgInTaskBar "The transaction date falls within a closed period", "Y", "N"
                 txtPostingDate.SetFocus
                 Exit Sub
              End If
      End If
      If szCallingForm = "frmDemands3" Then
         Label1(42).Caption = "Posting Date:"
         If frmDemands3.tabDmdRcpt.Tab = 0 Then
            If TextBoxFormatDate(txtPostingDate) Then
               frmDemands3.lblPostingDate.ToolTipText = txtPostingDate.text
               txtPostingDate.text = ""
               Unload Me
            End If
            frmDemands3.cmdaddnewline.SetFocus
         End If
         If frmDemands3.tabDmdRcpt.Tab = 2 Then
            If TextBoxFormatDate(txtPostingDate) Then
               frmDemands3.lblRptPostingDate.ToolTipText = txtPostingDate.text
               txtPostingDate.text = ""
               Unload Me
            End If
'Debug.Print frmDemands3.lblRptPostingDate.Visible
            frmDemands3.txtSPDate.SetFocus
         End If
      End If
    'Resolved by BOSL
    'Below line added by anol 29 Mar 2015
    'issue 549: Demand receipts not working note 3
     If szCallingForm = "frmReceiptEdit" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmReceiptEdit.lblRptPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If
      
      'End of modification
      If szCallingForm = "frmBatchDemands" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmBatchDemands.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If

      If szCallingForm = "frmBankTranEdit" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmBankTranEdit.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If

      If szCallingForm = "frmBankTransfer" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmBankTransfer.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If

      If szCallingForm = "frmBRPreForm" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmBRPreForm.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If

      If szCallingForm = "frmBPPreForm" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmBPPreForm.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If

      If szCallingForm = "frmNJ_Entry" Then
          Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmNJ_Entry.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If

      If szCallingForm = "frmPO2PI" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmPO2PI.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If
      'added by anol 23 Aug 2015
      If szCallingForm = "frmPaymentEdit" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmPaymentEdit.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If
      'end of addition
      'Anol 10 Jun 2015
       If szCallingForm = "frmBatchRpt" Then
         Label1(42).Caption = "Posting Date:"
         If TextBoxFormatDate(txtPostingDate) Then
            frmBatchRpt.lblPostingDate.ToolTipText = txtPostingDate.text
            txtPostingDate.text = ""
            Unload Me
         End If
      End If
      If szCallingForm = "frmPurchaseExpense" Then
         Label1(42).Caption = "Posting Date:"
         If frmPurchaseExpense.tabPurExp.Tab = 0 Then
            If TextBoxFormatDate(txtPostingDate) Then
               frmPurchaseExpense.lblPostingDate.ToolTipText = txtPostingDate.text
               txtPostingDate.text = ""
               Unload Me
            End If
         Else
            If TextBoxFormatDate(txtPostingDate) Then
               frmPurchaseExpense.lblPayPostingDate.ToolTipText = txtPostingDate.text
               txtPostingDate.text = ""
               Unload Me
            End If
         End If
      End If
      If szCallingForm = "frmRevPaymnent" Then
         Label1(42).Caption = "Allocation Date:"
         
      End If
   End If

   TextBoxKeyPrsDate txtPostingDate, KeyAscii
End Sub

Private Sub txtPostingDate_LostFocus()
   Unload Me
End Sub
