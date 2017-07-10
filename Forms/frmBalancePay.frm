VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmBalancePay 
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "frmBalancePay.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChange 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   14
      Top             =   4800
      Width           =   2055
   End
   Begin XPControls.XPButton cmdCheckIn 
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Check-In"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtAmountPaid 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   10
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtRemBal 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtAdvPay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtGName 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtBookNum 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin XPControls.XPButton cmdBack 
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Back"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "* put the amount given by the guest here and it must be higher than the remaining balance to checkin."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Payment Section"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   16
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmBalancePay.frx":455B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   5175
      Left            =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Balance:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Payment:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RF Number:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1200
   End
End
Attribute VB_Name = "frmBalancePay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdCheckIn_Click()
If frmCheckInBooked.txtTranType = 1 Then

        If MsgBox("Are you sure you want to check-in this guest? This process will not print OR", vbYesNo, "Confirm Check-In") = vbYes Then
        
        strSQL = "select * from tblbilling_booking where record_num=" & frmCheckInBooked.ListView1.SelectedItem.SubItems(11) & ""
        
        Set recSet = New ADODB.Recordset
        
        With recSet
        .Open strSQL, Conn, 3, 3
        
        !remaining_balance = 0
        !second_payment = txtRemBal.Text
        
        .Update
        
        End With
        
        strSQL = "select * from tblbooking_info where record_num=" & frmCheckInBooked.ListView1.SelectedItem.SubItems(11) & ""
        
        Set recSet = New ADODB.Recordset
        
        With recSet
        .Open strSQL, Conn, 3, 3
        !booking_status = 0
        !check_in_time = Format$(Now, "hh:mm AM/PM")
        
        .Update
        
        End With
        
        strSQL = "select * from tblroom_status where room_num=" & frmCheckInBooked.ListView1.SelectedItem.SubItems(2) & ""
        
        
        Set recSet = New ADODB.Recordset
        
        With recSet
        .Open strSQL, Conn, 3, 3
        
        !room_status = 1
        !customer_number = frmCheckInBooked.ListView1.SelectedItem.SubItems(5)
        .Update
        
        End With
        
        MsgBox "Check-In successfull. Returning into previous form.", vbOKOnly, "Check-In Success"
        
        Unload Me
        
        frmCheckInBooked.ListView1.ListItems.Clear
        frmCheckInBooked.ListView2.ListItems.Clear
        
        frmCheckInBooked.ListViewProp
        
        
        Else
        
        End If
        
ElseIf frmCheckInBooked.txtTranType = 2 Then
         
    If MsgBox("Are you sure you want to check-in this guest? This process will not print OR", vbYesNo, "Confirm Check-In") = vbYes Then
    
    
        MigrateReserve
        
        MsgBox "Check-In successfull. Returning into previous form.", vbOKOnly, "Check-In Success"
        
        Unload Me
        
        frmCheckInBooked.ListView1.ListItems.Clear
        frmCheckInBooked.ListView2.ListItems.Clear
        
        frmCheckInBooked.ListViewProp
        
    End If
End If
End Sub

Private Sub Form_Load()
DBCON
FormPos frmBalancePay

If frmCheckInBooked.txtTranType = 1 Then

GetBillingInfo

ElseIf frmCheckInBooked.txtTranType = 2 Then

GetReservationBilling

End If
cmdCheckIn.Enabled = False
End Sub

Private Sub GetReservationBilling()

txtBookNum.Text = frmCheckInBooked.ListView2.SelectedItem.SubItems(6)
txtGName.Text = frmCheckInBooked.ListView2.SelectedItem.SubItems(1)
txtAdvPay.Text = 0
txtRemBal.Text = frmCheckInBooked.ListView2.SelectedItem.SubItems(7)
txtTotal.Text = frmCheckInBooked.ListView2.SelectedItem.SubItems(7)

End Sub

Private Sub MigrateReserve()
Dim FName, LName, TranType As String
Dim MaxBookingRec, MaxBookRec, GNum, RecNum, ResNum, RoomNum, UserID, TotalPayment As Integer
Dim ResInDate, ResOutDate, DateDone, ExpectTimeIn, ExpectTimeOut As Date

strSQL = "select * from tblreservation_info where record_num=" & frmCheckInBooked.ListView2.SelectedItem.SubItems(9) & ""


Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

If Not .EOF Then

GNum = frmCheckInBooked.ListView2.SelectedItem.SubItems(5)
FName = frmCheckInBooked.ListView2.SelectedItem.SubItems(10)
LName = frmCheckInBooked.ListView2.SelectedItem.SubItems(11)
RoomNum = !room_num
UserID = !user_id
TotalPayment = !total_payment
ResInDate = !reservation_date
ResOutDate = !out_date
DateDone = !date_done
ExpectTimeIn = Format$(!expected_time_in, "hh:mm AM/PM")
ExpectTimeOut = Format$(!expected_time_out, "hh:mm AM/PM")
!reservation_status = 0

.Update


End If
End With

strSQL = "select * from tblbooking_record"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!customer_num = GNum
MaxBookRec = !booking_num

.Update

End With

strSQL = "select * from tblbooking_info"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!booking_num = MaxBookRec
!booking_date = Format$(ResInDate, "mm/dd/yyyy")
!out_date = Format(ResOutDate, "mm/dd/yyyy")
!room_num = RoomNum
!user_id = UserID
!date_done = Format$(DateDone, "mm/dd/yyyy")
!check_in_time = Format$(Now, "hh:mm AM/PM")
!booking_status = 0
!expected_time_in = Format$(ExpectTimeIn, "hh:mm AM/PM")
!expected_time_out = Format$(ExpectTimeOut, "hh:mm AM/PM")
MaxBookingRec = !record_num
!tran_type = "WI-" & MaxBookingRec

.Update

End With

strSQL = "select * from tblbilling_booking"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!record_num = MaxBookingRec
If txtAdvPay.Text = "" Then
!advance_payment = 0
Else
!advance_payment = txtAdvPay.Text
End If
If txtRemBal.Text = "" Then
!second_payment = 0

!remaining_balance = 0
Else
!second_payment = txtRemBal.Text
!remaining_balance = 0
End If
!others = 0

!total = TotalPayment
!refund = 0
!user_id = UserID
!date_done = Format(Now, "mm/dd/yyyy")
.Update

End With

strSQL = "select * from tbldatesandtime where transaction_type='" & frmCheckInBooked.ListView2.SelectedItem.SubItems(6) & "'"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.Delete
.Update
End With

strSQL = "select * from tbldatesandtime"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!transaction_type = "WI-" & MaxBookingRec
!date_check_in = ResInDate
!date_check_out = ResOutDate
!check_in_time = ExpectTimeIn
!check_out_time = ExpectTimeOut
!room_number = RoomNum

.Update

End With

strSQL = "select * from tblroom_status where room_num=" & RoomNum & ""

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

!room_status = 1
!customer_number = GNum

.Update

.Close

End With

End Sub


Private Sub GetBillingInfo()

strSQL = "select * from tblbilling_booking where record_num=" & frmCheckInBooked.ListView1.SelectedItem.SubItems(11) & ""

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

If Not .EOF Then

txtBookNum.Text = "BK-" & !record_num
txtGName.Text = frmCheckInBooked.ListView1.SelectedItem.SubItems(1)
txtAdvPay.Text = !advance_payment
txtRemBal.Text = !remaining_balance
txtTotal.Text = !total


End If

.Close

End With
End Sub

Private Sub txtAmountPaid_Change()
If Val(txtAmountPaid.Text) >= Val(txtRemBal.Text) Then

cmdCheckIn.Enabled = True
txtChange.Text = Val(txtAmountPaid.Text) - Val(txtRemBal.Text)

Else
cmdCheckIn.Enabled = False
txtChange.Text = ""
End If

End Sub
