VERSION 5.00
Begin VB.Form frmBilling 
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRFNumHidden 
      Height          =   375
      Left            =   1440
      TabIndex        =   32
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtOthers 
      Height          =   495
      Left            =   5160
      TabIndex        =   31
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtBalance 
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtNumberOfDays 
      Height          =   495
      Left            =   1560
      TabIndex        =   27
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtCheckInTime 
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintOR 
      Caption         =   "&Print OR"
      Height          =   495
      Left            =   5760
      TabIndex        =   23
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtAmountPay 
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Height          =   495
      Left            =   5160
      TabIndex        =   21
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtRoomRate 
      Height          =   495
      Left            =   5160
      TabIndex        =   20
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtRoomNum 
      Height          =   495
      Left            =   5160
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtRoomType 
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtDateOut 
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtBookReserveDate 
      Height          =   495
      Left            =   1560
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtTotalPax 
      Height          =   495
      Left            =   1560
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtGName 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtGnumber 
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtRFNum 
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Others:"
      Height          =   495
      Left            =   3720
      TabIndex        =   30
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblBalance 
      Caption         =   "Remaining Balance:"
      Height          =   495
      Left            =   3720
      TabIndex        =   28
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblNumberOfDays 
      Caption         =   "Number Of Days:"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblCheckInTime 
      Caption         =   "Check-In Time:"
      Height          =   495
      Left            =   3720
      TabIndex        =   24
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Amount To Pay:"
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Total:"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Guest Name:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Room Rate:"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Total Pax:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Date Out:"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblBookReserveDate 
      Caption         =   "Reservation Date:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Room Number:"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Room Type:"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Guest Number:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblBookReserveNum 
      Caption         =   "RF Num:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Billing"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrintOR_Click()
If frmReservation.cmdOkay.Tag = 1 Or frmReservation.cmdOkay.Tag = 3 Then

Billing

ElseIf frmReservation.cmdOkay.Tag = 2 Then

MsgBox "Reservation Success"

Unload Me
Unload frmReservation

End If

MsgBox "Print Success,Returning to main form...", vbInformation + vbOKOnly, "Print Success"

Unload Me
Unload frmReservation
End Sub

Private Sub Form_Load()
DBCON

txtRFnum.Locked = True
txtGNumber.Locked = True
txtGName.Locked = True
txtTotalPax.Locked = True
txtBookReserveDate.Locked = True
txtDateOut.Locked = True
txtRoomType.Locked = True
txtRoomNum.Locked = True
txtRoomRate.Locked = True
txtTotal.Locked = True
txtAmountPay.Locked = True
txtCheckInTime.Locked = True

lblCheckInTime.Visible = False
txtCheckInTime.Visible = False
lblBalance.Visible = False
txtBalance.Visible = False

txtGNumber.Text = frmReservation.txtGNumber.Text
txtGName.Text = frmReservation.txtGName.Text
txtTotalPax.Text = frmReservation.txtHeadCount.Text
txtBookReserveDate.Text = frmReservation.txtsDate.Text
txtDateOut.Text = frmReservation.txtDateOut.Text
txtRoomType.Text = frmReservation.cboRoomType.Text
txtRoomNum.Text = frmReservation.cboRoomNumber.Text
txtRoomRate.Text = frmReservation.txtRoomRate
txtNumberOfDays.Text = frmReservation.txtNumberOfDays.Text
txtOthers.Text = Val(frmReservation.txtExtendHour.Text) * 100

    If Val(frmReservation.txtNumberOfDays.Text) = 0 Then ' if the guest will check-in and check-out on same date
    
        txtTotal.Text = 1 * Val(frmReservation.txtRoomRate.Text)
        
    Else
    
        txtTotal.Text = (Val(frmReservation.txtNumberOfDays.Text) * Val(frmReservation.txtRoomRate.Text) + Val(txtOthers.Text))
        
    End If
    
    
    If frmReservation.cmdOkay.Tag = 1 Then
    
     strSQL = "select max(record_num) as MaxBooknum from tblbooking_info"
     
      Set recSet = New ADODB.Recordset
      
       With recSet
       .Open strSQL, Conn, 3, 3
       
       txtRFnum.Text = "BK" & "-" & !maxbooknum
       txtRFNumHidden.Text = !maxbooknum
       
       .Close
       
       End With
       
       txtAmountPay.Text = Val(txtTotal.Text) / 2
       txtBalance.Text = Val(txtTotal.Text) - Val(txtAmountPay.Text)
    
    ElseIf frmReservation.cmdOkay.Tag = 2 Then
    
    strSQL = "select max(record_num) as MaxReservationnum from tblreservation_info"
    
     Set recSet = New ADODB.Recordset
     
      With recSet
      .Open strSQL, Conn, 3, 3
      
      txtRFnum.Text = "RV" & "-" & !maxreservationnum
      txtRFNumHidden.Text = !maxreservationnum
      
      .Close
      End With
      
      txtAmountPay.Text = 0
      
    ElseIf frmReservation.cmdOkay.Tag = 3 Then
    
     strSQL = "select max(record_num) as MaxWalkInNum from tblbooking_info"
     
      Set recSet = New ADODB.Recordset
      
       With recSet
       
        .Open strSQL, Conn, 3, 3
        
         txtRFnum.Text = "WI" & "-" & !maxwalkinnum
         txtRFNumHidden.Text = !maxwalkinnum
        
        .Close
       
       End With
       
       txtBalance.Text = 0
       txtAmountPay.Text = Val(txtRoomRate.Text) * Val(txtNumberOfDays.Text)
    
    End If
    
End Sub

Sub Billing()

strSQL = "select * from tblbilling_booking"
         
       Set recSet = New ADODB.Recordset
       
        With recSet
        .Open strSQL, Conn, 3, 3
        
        .AddNew
        !record_num = txtRFNumHidden.Text
        !advance_payment = txtAmountPay.Text
        !remaining_balance = txtBalance.Text
        !total = txtTotal.Text
        !refund = Val(txtAmountPay.Text) / 2
        !user_id = frmMain.txtUserID.Text
        !date_done = Format(Now, "mm/dd/yyyy")
        If txtOthers.Text = "" Then
        !others = 0
        Else
        !others = txtOthers.Text
        End If
        
        .Update
        .Close
        
        End With

End Sub
