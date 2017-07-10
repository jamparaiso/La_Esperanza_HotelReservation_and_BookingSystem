VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmCheckInBooked 
   ClientHeight    =   4635
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "frmCheckInBooked.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTranType 
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Top             =   4800
      Width           =   735
   End
   Begin XPControls.XPButton cmdBack 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Back"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjXTab.XTab XTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6588
      TabCaption(0)   =   "Tab 0"
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "ListView1"
      TabCaption(1)   =   "Tab 1"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "ListView2"
      TabCaption(2)   =   "Tab 2"
      ActiveTab       =   1
      ActiveTabBackStartColor=   12648447
      ActiveTabBackEndColor=   12648384
      InActiveTabBackStartColor=   16761087
      InActiveTabBackEndColor=   12640511
      ActiveTabForeColor=   255
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   16744576
      BottomRightInnerBorderColor=   16744576
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      PictureMaskColor=   16777215
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imglstIcons1"
         SmallIcons      =   "imglstIcons1"
         ColHdrIcons     =   "imglstIcons1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList imglstIcons1 
      Left            =   600
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":455B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":4DAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":2E9D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":585F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":58B8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":5E7AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":883D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":8F8D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":BA5AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":BAE87
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":BB761
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":BC03B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":BC1C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckInBooked.frx":E5DE9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "popUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCheck 
         Caption         =   "Check-In"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookInfo 
         Caption         =   "Booking Information"
      End
   End
   Begin VB.Menu mnuPopReserve 
      Caption         =   "popUpMenuReserve"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckInReserve 
         Caption         =   "Check-In"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReseveInfo 
         Caption         =   "Reservation Information"
      End
   End
End
Attribute VB_Name = "frmCheckInBooked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objList1 As ListItem
Dim objList2 As ListItem
Dim CustomerNumber, RoomNumber As Integer

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub Form_Load()
FormPos frmCheckInBooked
DBCON
TabProp
ListViewProp

End Sub


Sub ListViewProp()

Set ListView1.Icons = imglstIcons1
Set ListView2.Icons = imglstIcons1

With ListView1

.Arrange = lvwAutoTop
.HotTracking = True
.LabelEdit = lvwManual
.BorderStyle = ccFixedSingle
.Appearance = ccFlat
.OLEDragMode = ccOLEDragAutomatic

For a = 1 To 12
.ColumnHeaders.Add a
Next a

End With

With ListView2

.Arrange = lvwAutoTop
.HotTracking = True
.LabelEdit = lvwManual
.BorderStyle = ccFixedSingle
.Appearance = ccFlat
.OLEDragMode = ccOLEDragAutomatic

For c = 1 To 12
.ColumnHeaders.Add c
Next c

End With


'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'for booking

strSQL = "SELECT tblBooking_Info.booking_date, tblBooking_Info.out_date, " & _
         "tblBooking_Info.room_num, tblBooking_Info.booking_status, tblCustomer_Info.customer_fname, " & _
         "tblCustomer_Info.customer_lname, tblBooking_Info.record_num, tblBooking_Info.expected_time_in, " & _
         "tblBooking_Info.expected_time_out, tblBilling_Booking.advance_payment, " & _
         "tblBilling_Booking.remaining_balance, tblBilling_Booking.total, " & _
         "tblBilling_Booking.refund, tblCustomer_Info.customer_num, tblBooking_Info.user_id " & _
         "FROM (tblRoom_Status INNER JOIN ((tblCustomer_Info INNER JOIN " & _
         "tblBooking_Record ON tblCustomer_Info.customer_num = tblBooking_Record.customer_num) " & _
         "INNER JOIN tblBooking_Info ON " & _
         "tblBooking_Record.booking_num = tblBooking_Info.booking_num) ON " & _
         "tblRoom_Status.room_num = tblBooking_Info.room_num) INNER JOIN " & _
         "tblBilling_Booking ON tblBooking_Info.record_num = tblBilling_Booking.record_num " & _
         "WHERE (((tblBooking_Info.booking_date)=Date()) AND " & _
         "((tblBooking_Info.booking_status)=2));"




Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

Do While Not .EOF

 i = 1
 
 Set objList1 = ListView1.ListItems.Add(i, , "BK-" & !record_num & " Room Number: " & !room_num & " Customer Name: " & !customer_fname & " " & !customer_lname, 2)
     objList1.SubItems(1) = !customer_fname & " " & !customer_lname
     objList1.SubItems(2) = !room_num
     objList1.SubItems(3) = !booking_date
     objList1.SubItems(4) = !out_Date
     objList1.SubItems(5) = !customer_num
     objList1.SubItems(6) = "BK-" & !record_num
     objList1.SubItems(7) = !advance_payment
     objList1.SubItems(8) = !remaining_balance
     objList1.SubItems(9) = !total
     objList1.SubItems(10) = !user_id
     objList1.SubItems(11) = !record_num
     
 i = i + 1

.MoveNext

Loop

End With

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'for reservation

strSQL = "SELECT tblReservation_Info.reservation_date, tblReservation_Info.out_date, " & _
         "tblReservation_Info.room_num, tblReservation_Info.reservation_status, " & _
         "tblCustomer_Info.customer_fname, tblCustomer_Info.customer_lname, " & _
         "tblReservation_Info.total_payment, tblReservation_Info.tran_type, " & _
         "tblReservation_Info.reservation_num, tblReservation_Info.record_num, " & _
         "tblReservation_Record.customer_num, tblReservation_Info.user_id " & _
         "FROM tblRoom_Status INNER JOIN ((tblCustomer_Info INNER JOIN tblReservation_Record ON " & _
         "tblCustomer_Info.customer_num = tblReservation_Record.customer_num) INNER JOIN " & _
         "tblReservation_Info ON " & _
         "tblReservation_Record.reservation_num = tblReservation_Info.reservation_num) ON " & _
         "tblRoom_Status.room_num = tblReservation_Info.room_num " & _
         "WHERE (((tblReservation_Info.reservation_date)=Date()) AND " & _
         "((tblReservation_Info.reservation_status)=1));"

         
 Set recSet = New ADODB.Recordset
 
 With recSet
 .Open strSQL, Conn, 3, 2
 
 Do While Not .EOF
 
 c = 1
 
 Set objList2 = ListView2.ListItems.Add(c, , "RV-" & !record_num & " Room Number:" & !room_num & " Customer Name:" & !customer_fname & " " & !customer_lname, 2)
     objList2.SubItems(1) = !customer_fname & " " & !customer_lname
     objList2.SubItems(2) = !room_num
     objList2.SubItems(3) = !reservation_date
     objList2.SubItems(4) = !out_Date
     objList2.SubItems(5) = !customer_num
     objList2.SubItems(6) = !tran_type
     objList2.SubItems(7) = !total_payment
     objList2.SubItems(8) = !user_id
     objList2.SubItems(9) = !record_num
     objList2.SubItems(10) = !customer_fname
     objList2.SubItems(11) = !customer_lname
     
     
 c = c + 1
 
 .MoveNext
 Loop
 
 End With
 
End Sub

Private Sub mnuReseveInfo_Click()
On Error GoTo errcatch

With frmTranInfo
.txtBookRFNum = ListView2.SelectedItem.SubItems(6)
.txtBookGName = ListView2.SelectedItem.SubItems(1)
.txtBookCheckInDate = ListView2.SelectedItem.SubItems(3)
.txtBookCheckOutDate = ListView2.SelectedItem.SubItems(4)
.txtBookAdvPay = 0
.txtBookRemBal = ListView2.SelectedItem.SubItems(7)
.txtBookTotal = ListView2.SelectedItem.SubItems(7)
.txtDoneBy = ListView2.SelectedItem.SubItems(8)

.Show vbModal, frmCheckInBooked
End With

errcatch:
Select Case Err.Number

Case 91

MsgBox "No Records Found. No Bookings are listed today.", vbInformation + vbOKOnly, "No Records"

Exit Sub

End Select
End Sub

Sub TabProp()

With XTab1
.TabCount = 2
.TabCaption(0) = "Booking"
.TabCaption(1) = "Reservation"


End With
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errcatch

    If Button = vbRightButton Then
    PopupMenu mnuPopUp
    End If
    
errcatch:
Select Case Err.Number

Case 91

MsgBox "No Records Found. No Bookings are listed today.", vbInformation + vbOKOnly, "No Records"

Exit Sub

End Select
End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
PopupMenu mnuPopReserve

End If
End Sub

Private Sub mnuBookInfo_Click()
On Error GoTo errcatch

With frmTranInfo
.txtBookRFNum = ListView1.SelectedItem.SubItems(6)
.txtBookGName = ListView1.SelectedItem.SubItems(1)
.txtBookCheckInDate = ListView1.SelectedItem.SubItems(3)
.txtBookCheckOutDate = ListView1.SelectedItem.SubItems(4)
.txtBookAdvPay = ListView1.SelectedItem.SubItems(7)
.txtBookRemBal = ListView1.SelectedItem.SubItems(8)
.txtBookTotal = ListView1.SelectedItem.SubItems(9)
.txtDoneBy = ListView1.SelectedItem.SubItems(10)

.Show vbModal, frmCheckInBooked
End With


errcatch:
Select Case Err.Number

Case 91

MsgBox "No Records Found. No Bookings are listed today.", vbInformation + vbOKOnly, "No Records"

Exit Sub

End Select
End Sub

Private Sub mnuCheck_Click()
On Error GoTo errcatch

MsgBox "Checking if the guest have unpaid balance....", vbOKOnly, "Checking..."

strSQL = "select * from tblbilling_booking where record_num=" & ListView1.SelectedItem.SubItems(11) & ""

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

If Not .EOF Then

    If Val(!remaining_balance) > 0 Then
    MsgBox "The guest have remaining balance. Press Ok to proceed in payment form.", vbOKCancel, "Balance Check"
    
    txtTranType.Text = 1
    frmBalancePay.Show vbModal, frmCheckInBooked
    
    ElseIf Val(!remaining_balance) = 0 Then
    
    If MsgBox("The guest don't have any unpaid balance. Press Ok to Check-In the guest.", vbOKCancel, "Balance Check") = vbOK Then
    
    CheckInTheGuest
    
    MsgBox "The guest has been successfully checked-in.", vbOKOnly, "Check-In Success"
    
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    
    ListViewProp
    
    End If
    
    Exit Sub
    
    End If

End If

End With


errcatch:
Select Case Err.Number

Case 91

MsgBox "No Records Found. No Bookings are listed today.", vbInformation + vbOKOnly, "No Records"

Exit Sub

End Select
End Sub

Private Sub CheckInTheGuest()

strSQL = "SELECT tblBooking_Info.booking_date, tblBooking_Info.out_date, tblBooking_Info.room_num, " & _
         "tblBooking_Info.booking_status, tblCustomer_Info.customer_fname, " & _
         "tblCustomer_Info.customer_lname, tblBooking_Info.record_num, " & _
         "tblBooking_Info.expected_time_in, tblBooking_Info.expected_time_out, " & _
         "tblBilling_Booking.advance_payment, tblBilling_Booking.remaining_balance, " & _
         "tblBilling_Booking.total, tblBilling_Booking.refund, tblCustomer_Info.customer_num, " & _
         "tblBooking_Info.user_id, tblBooking_Info.check_in_time " & _
         "FROM tblRoom_Status INNER JOIN ((tblCustomer_Info INNER JOIN tblBooking_Record ON " & _
         "tblCustomer_Info.customer_num = tblBooking_Record.customer_num) INNER JOIN " & _
         "(tblBooking_Info INNER JOIN tblBilling_Booking ON " & _
         "tblBooking_Info.record_num = tblBilling_Booking.record_num) ON " & _
         "tblBooking_Record.booking_num = tblBooking_Info.booking_num) ON " & _
         "tblRoom_Status.room_num = tblBooking_Info.room_num " & _
         "WHERE (((tblBooking_Info.booking_date)=Date()) AND " & _
         "((tblBooking_Info.booking_status)=2) AND " & _
         "((tblBooking_Info.record_num)=" & ListView1.SelectedItem.SubItems(11) & "));"


Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2
If Not .EOF Then
!booking_status = 0
CustomerNumber = !customer_num
RoomNumber = !room_num
!check_in_time = Format$(Now, "hh:mm AM/PM")

.Update
End If
End With

strSQL = "select * from tblroom_status where room_num=" & RoomNumber & ""


Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

If Not .EOF Then

!room_status = 1
!customer_number = CustomerNumber

.Update
.Close
End If
End With
End Sub

Private Sub mnuCheckInReserve_Click()
On Error GoTo errcatch

If MsgBox("Are you sure you want to check-in this guest? Press Ok to continue in billing form.", vbOKCancel, "Confirm Check-In") = vbOK Then

txtTranType.Text = 2

frmBalancePay.Show vbModal, frmCheckInBooked

End If



errcatch:
Select Case Err.Number

Case 91

MsgBox "No Records Found. No Bookings are listed today.", vbInformation + vbOKOnly, "No Records"

Exit Sub

End Select
End Sub


