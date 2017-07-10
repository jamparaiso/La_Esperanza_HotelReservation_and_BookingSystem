VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReservation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservation"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14310
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   14310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExtendHour 
      Height          =   495
      Left            =   1560
      TabIndex        =   35
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtCheckInTime 
      Height          =   495
      Left            =   5160
      TabIndex        =   33
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtDateOut 
      Height          =   495
      Left            =   8640
      TabIndex        =   31
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtcmdOkayTag 
      Height          =   495
      Left            =   12960
      TabIndex        =   30
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtRoomTypeIndex 
      Height          =   495
      Left            =   11520
      TabIndex        =   28
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtMonthNum 
      Height          =   495
      Left            =   10080
      TabIndex        =   27
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtRoomRate 
      Height          =   495
      Left            =   5160
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtMaxPax 
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtsDate 
      Height          =   495
      Left            =   8640
      TabIndex        =   20
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtfDate 
      Height          =   495
      Left            =   7200
      TabIndex        =   19
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ComboBox cboRoomNumber 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox cboRoomType 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "cmdOkay"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   5400
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   64159745
      CurrentDate     =   41153
   End
   Begin VB.TextBox txtNumberOfDays 
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtHeadCount 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtGName 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtGNumber 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   64159745
      CurrentDate     =   41152
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   6960
      TabIndex        =   29
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label13 
      Caption         =   "100/Hour"
      Height          =   495
      Left            =   2760
      TabIndex        =   36
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Extend Hour:"
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblCheckInTime 
      Caption         =   "Check-In Time"
      Height          =   495
      Left            =   3720
      TabIndex        =   32
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "*room rate may vary depeding on the peak season"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   6360
      Width           =   3555
   End
   Begin VB.Label Label11 
      Caption         =   "*"
      Height          =   495
      Left            =   6480
      TabIndex        =   25
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Room Rate:"
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Room Max Pax:"
      Height          =   495
      Left            =   3720
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Number of Days:"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Room Number:"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Date Out:"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblBookReserve 
      Caption         =   "Reservation date:"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Room Type:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Guest Pax:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Guest Number:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fDate As Date
Dim sdate As Date
Dim roomIndex As Integer
Dim BookingNum As Integer
Dim ReservationNum As Integer
Dim WalkInNum As Integer
Dim Access As Integer

Private Sub cboRoomType_Click()
txtRoomTypeIndex.Text = cboRoomType.ListIndex 'this returns what is the current index of the cboroomtype
cboRoomNumber.Clear 'clear the cboroomnumver prior to change in roomtype
cboRoom_Number 'gets the current available rooms
End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub

Sub ChangeRoomStatus()

 strSQL = "select * from tblroom_status where room_num=" & cboRoomNumber.Text & ""
 
  Set recSet = New ADODB.Recordset
  
   With recSet
   .Open strSQL, Conn, 3, 2
   
    !room_status = 1
    
    .Update
   
   .Close
   
   End With

End Sub
Private Sub VerifyRooms()

strSQL = "select * from tblbooking_info"

Set recSet = New ADODB.Recordset

 With recSet
 .Open strSQL, Conn, 3, 3
 
    Do While .EOF = False
    
        If DTPicker1.Value >= !booking_date And DTPicker1.Value <= !out_date And cboRoomNumber.Text Then
        
        MsgBox "The Room is already Booked/Reserved on this date. Please Choose Another", vbInformation + vbOKOnly, "Room Not Available"
            cboRoomNumber.SetFocus
            Access = 0
            Exit Sub
        ElseIf DTPicker2.Value >= !booking_date And DTPicker2.Value <= !out_date And cboRoomNumber.Text Then
        MsgBox "The Room is already Booked/Reserved on this date. Please Choose Another", vbInformation + vbOKOnly, "Room Not Available"
            cboRoomNumber.SetFocus
            Access = 0
            Exit Sub
        ElseIf DTPicker1.Value <= !booking_date And DTPicker2.Value >= !out_date And cboRoomNumber.Text Then
        
        MsgBox "The Room is already Booked/Reserved on this date. Please Choose Another", vbInformation + vbOKOnly, "Room Not Available"
            cboRoomNumber.SetFocus
            Access = 0
            Exit Sub
            
        Else
        
        Access = 1
        End If
        
       .MoveNext
 
    Loop
 End With

End Sub
Private Sub cmdOkay_Click()
txtcmdOkayTag.Text = cmdOkay.Tag
If txtGName.Text <> "" And txtHeadCount.Text <= txtMaxPax.Text Then
VerifyRooms

If Access = 1 Then '1

If txtGNumber.Text <> "" Then

  If Val(txtHeadCount.Text) <= Val(txtMaxPax.Text) Then

    MsgBox "Making Billing Statement....Press Okay to continue", vbOKOnly, "Please Wait"
 
        If cmdOkay.Tag = 1 Then

                NewBookingRecord
 
                NewBookingInfo
 
                frmBilling.lblBookReserveDate.Caption = "Booking Date"
                frmBilling.lblBalance.Visible = True
                frmBilling.txtBalance.Visible = True
 
        ElseIf cmdOkay.Tag = 2 Then

                NewReservationRecord
 
                NewReservationInfo
 
        ElseIf cmdOkay.Tag = 3 Then

                NewWalkInRecord

                NewWalkInInfo

                ChangeRoomStatus

                frmBilling.lblCheckInTime.Visible = True
                frmBilling.txtCheckInTime.Visible = True
                frmBilling.txtCheckInTime.Text = txtCheckInTime.Text

                frmBilling.lblBookReserveDate.Caption = "Walk-In Date"

                End If

                frmBilling.Show vbModal, frmReservation
                
    Else
                
      MsgBox "The Room cannot accomodate " & txtHeadCount.Text & " people", vbCritical + vbOKOnly, "Room Occupants Exceeded"
                
      txtHeadCount.Text = ""
                
    End If

Else

End If
End If

Else

MsgBox "No Records Found", vbCritical + vbOKOnly, "Error"

End If
End Sub

Private Sub NewReservationInfo()

strSQL = "select * from tblreservation_info"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   .AddNew
   
    !reservation_num = ReservationNum 'this holds the max record_num on tblreservation_record,initializes on sub NewReservationRecord
    !room_num = cboRoomNumber.Text
    !reservation_date = DTPicker1.Value
    !out_date = DTPicker2.Value
    !user_id = frmMain.txtUserID.Text
    !date_done = txtfDate.Text
    !reservation_status = 1
    !extend_hour = txtExtendHour.Text
    
    
    .Update
    
  
  .Close
  
  End With

End Sub

Private Sub NewBookingInfo()

'this will create a new booking_info
'this will need the BookingNum variable

strSQL = "select * from tblBooking_info"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   .AddNew
   
   !booking_num = BookingNum 'this holds the max record_num on tblbook_record,initializes on sub NewBookingRecord
   !booking_date = DTPicker1.Value
   !out_date = DTPicker2.Value
   !room_num = cboRoomNumber.Text
   !user_id = frmMain.txtUserID.Text 'currently log in user id
   !date_done = txtfDate.Text
   !booking_status = 1
   !extend_hour = txtExtendHour.Text
   
   .Update
  
  .Close
  
  End With

End Sub

Sub NewWalkInInfo()

'this will create a record on tblbooking_info
'this will need the WalkInNum variable

strSQL = "select * from tblBooking_info"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   .AddNew
   
   !booking_num = WalkInNum 'this holds the max record_num on tblbook_record,initializes on sub NewBookingRecord
   !booking_date = DTPicker1.Value
   !out_date = DTPicker2.Value
   !room_num = cboRoomNumber.Text
   !user_id = frmMain.txtUserID.Text 'currently log in user id
   !date_done = txtfDate.Text
   !check_in_time = txtCheckInTime.Text
   !booking_status = 1
   !extend_hour = txtExtendHour.Text
   
   .Update
  
  .Close
  
  End With

End Sub

Private Sub DTPicker1_Change()
txtsDate.Text = DTPicker1.Value

DTPicker2.Value = DTPicker1.Value + 1 'this code prevents the check out date become lower than check in date

txtNumberOfDays.Text = ""
txtNumberOfDays.Text = DTPicker2.Value - DTPicker1.Value 'computes the dates difference

txtDateOut.Text = DTPicker2.Value

cboRoomType.Clear
cboRoom_Types
cboRoomNumber.Clear
cboRoom_Number

End Sub

Private Sub DTPicker2_Change()

If DTPicker2.Value < DTPicker1.Value Then 'if the date on datepicker2 is lower than datepicker1

 txtNumberOfDays.Text = 0
 
 MsgBox "Check-out Date is lower than" & " " & frmReservation.Caption & " date!Please Change the Checkout date", vbCritical + vbOKOnly, "Error"

 DTPicker2.Value = DTPicker1.Value + 1 'value if datepicker2 +1
 
 txtNumberOfDays.Text = DTPicker2.Value - DTPicker1.Value 'computes the dates difference
 
 txtDateOut.Text = DTPicker2.Value
 
Else 'if the datepicker2 is greater than datepicker1

txtNumberOfDays.Text = DTPicker2.Value - DTPicker1.Value 'computes the dates difference

txtDateOut.Text = DTPicker2.Value

End If

cboRoomType.Clear
cboRoom_Types
cboRoomNumber.Clear
cboRoom_Number
End Sub

Sub NewReservationRecord()

'this creates new record on tblreservation_record

strSQL = "select * from tblreservation_record"

 Set recSet = New ADODB.Recordset
 
  With recSet
   
   .Open strSQL, Conn, 3, 3
   
    .AddNew
    
     !customer_num = txtGNumber.Text
     
     .Update
   
   .Close
   
  End With
  
'this will get the max reservation_num that will be needed to new reservation_info

strSQL = "select max(reservation_num) as MaxReservationNum from tblreservation_record"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   ReservationNum = !maxreservationnum 'needed on tblreservation_info and on sub NewReservationInfo
  
  .Close
  
  End With
  
End Sub

Sub NewBookingRecord()

strSQL = "select * from tblbooking_record"

 Set recSet = New ADODB.Recordset
 
  With recSet
   
   .Open strSQL, Conn, 3, 3
   
    .AddNew
    
     !customer_num = txtGNumber.Text
     
    .Update
   
   .Close
   
  End With
  
'this will get the max reservation_num that will be needed to new reservation_info

strSQL = "select max(booking_num) as MaxBookingNum from tblbooking_record"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   BookingNum = !maxbookingnum ' needed on tblbooking_info and on sub NewBookingInfo
  
  .Close
  
  End With
  
End Sub


Sub NewWalkInRecord()

strSQL = "select * from tblbooking_record"

 Set recSet = New ADODB.Recordset
 
  With recSet
   
   .Open strSQL, Conn, 3, 3
   
    .AddNew
    
     !customer_num = txtGNumber.Text
     
    .Update
   
   .Close
   
  End With
  
'this will get the max record_num that will be needed on tblbooking_info

strSQL = "select max(booking_num) as MaxWalkInNum from tblbooking_record"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   WalkInNum = !maxwalkinnum ' needed on tblbooking_info and on sub NewBookingInfo
  
  .Close
  
  End With
  
End Sub

Private Sub Form_Load()
DBCON

lblCheckInTime.Visible = False 'this is only needed on walk-in
txtCheckInTime.Visible = False

ListView_Prop

BookingList

cboRoom_Types 'initialize and populate cboroomtype

txtNumberOfDays.Locked = True
txtRoomRate.Locked = True
txtMaxPax.Locked = True
txtGName.Locked = True
txtExtendHour.Text = 0
txtHeadCount.Text = 1
txtHeadCount.MaxLength = 1


DTPicker1.Value = Format$(Now, "mm/dd/yyyy") 'this ensures that the date on dtpickers are the current
DTPicker2.Value = Format$(Now + 1, "mm/dd/yyyy")

fDate = DTPicker1.Value 'current date
txtfDate.Text = fDate 'variable content doesnt change

txtsDate.Text = DTPicker1.Value
txtDateOut.Text = DTPicker2.Value

txtMonthNum.Text = Month(Now) 'returns what month we are right now(in integer) this is used on the seasonal rates

txtNumberOfDays.Locked = True 'date difference between dtpicker1 and dtpicker2

txtNumberOfDays.Text = DTPicker2.Value - DTPicker1.Value 'computes the dates difference

End Sub




Private Sub txtExtendHour_Change()
If txtExtendHour.Text = "" Then
txtExtendHour.Text = 0
Else
End If
End Sub

Private Sub txtExtendHour_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtExtendHour)
End Sub

Private Sub txtGNumber_Change()

'to display the name of the customer
'-------------------------------------------------------------------------------------------------------
If txtGNumber.Text <> "" Then '0

    strSQL = "select * from tblcustomer_info where customer_num=" & txtGNumber.Text & ""

        Set recSet = New ADODB.Recordset
 
            With recSet
            
                .Open strSQL, Conn, 3, 2
    
                    If .EOF Or .BOF Then '1
     
                        txtGName.Text = ""
     
                    Else '1
    
                        txtGName.Text = !customer_fname & " " & !customer_lname
     
                    End If '1
     
                .Close

            End With
   
Else '0

txtGName.Text = ""

End If '0

End Sub

Sub cboRoom_Types()

'for cboroomtypes
'get all the list of rooms in db and assign a index that to be used in cboroomnumbeer
'---------------------------------------------------------------------------------------------------
strSQL = "select * from tblroom_info"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 2
  
   i = 0
  
   Do Until .EOF
   
    cboRoomType.AddItem !room_name, i 'i is the index
    
    .MoveNext
    
    i = i + 1
   
   Loop
  
  .Close
  
  End With
 
End Sub

Sub cboRoom_Number()

Select Case txtRoomTypeIndex.Text

Case 0

roomIndex = 1

Case 1

roomIndex = 2

Case 2

roomIndex = 3

Case 3

roomIndex = 4

Case 4

roomIndex = 5

End Select

'available rooms list
'------------------------------------------------------------------------------------------------------
If txtfDate.Text = DTPicker1.Value Then 'if the dtpicker1 date and current date match it means that it only list the available rooms

strSQL = "select * from tblroom_status where room_type=" & roomIndex & " and room_status= 0"


 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   Do Until .EOF
   
   cboRoomNumber.AddItem !room_num
   
   .MoveNext
   
   Loop
  
  .Close
  
  End With
  
Else 'if the date on dtpicker1 doesnt match on current date, all the rooms will be listed

strSQL = "select * from tblroom_status where room_type=" & roomIndex & ""

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   Do Until .EOF
   
   cboRoomNumber.AddItem !room_num
   
   .MoveNext
   
   Loop
  
  .Close
  
  End With

End If


  
'for rates, pax, etc,,,
'-----------------------------------------------------------------------------------------------------

 If txtMonthNum.Text = 1 Then 'if the month is january(prior to ati-atihan)the room rates will change
 
  strSQL = "select * from tblroom_info where room_type_num= " & roomIndex & " "
  
   Set recSet = New ADODB.Recordset
   
    With recSet
     .Open strSQL, Conn, 3, 2
     
      txtRoomRate.Text = !room_rate_on_season
      txtMaxPax.Text = !room_max_pax
     
     .Close
     
    End With


 Else 'if not january then normal rate are applied

    strSQL = "select * from tblroom_info where room_type_num= " & roomIndex & " "

        Set recSet = New ADODB.Recordset
 
            With recSet
  
                .Open strSQL, Conn, 3, 2
  
                    txtRoomRate.Text = !room_rate
                    txtMaxPax.Text = !room_max_pax
  
                .Close
  
            End With

 End If
 
End Sub

Sub ListView_Prop()

 With ListView1
 
  .View = lvwReport
  .FullRowSelect = True
  .GridLines = True
  .ColumnHeaders.Clear
  .ColumnHeaders.Add 1, , "Type", .Width * 0.15
  .ColumnHeaders.Add 2, , "Check-In Date", .Width * 0.18
  .ColumnHeaders.Add 3, , "Check-Out Date", .Width * 0.19
  .ColumnHeaders.Add 4, , "Room Number", .Width * 0.18
  .ColumnHeaders.Add 5, , "Guest Name", .Width * 0.3
 
 End With

End Sub

Sub BookingList()

Dim objList1 As ListItem
Dim objList2 As ListItem
Dim objList3 As ListItem
Dim objList4 As ListItem

'for ongoing book/walk-in

strSQL = "SELECT tblBooking_Info.booking_date, tblBooking_Info.out_date, tblBooking_Info.room_num, " & _
         "tblCustomer_Info.customer_fname, tblCustomer_Info.customer_lname, tblRoom_Status.room_status, " & _
         "tblBooking_Info.booking_status " & _
         "FROM tblRoom_Status INNER JOIN ((tblCustomer_Info INNER JOIN tblBooking_Record ON " & _
         "tblCustomer_Info.customer_num = tblBooking_Record.customer_num) INNER JOIN tblBooking_Info ON " & _
         "tblBooking_Record.booking_num = tblBooking_Info.booking_num) ON " & _
         "tblRoom_Status.room_num = tblBooking_Info.room_num " & _
         "WHERE (((tblRoom_Status.room_status)=1) AND ((tblBooking_Info.booking_status)=0));"

         
     Set recSet = New ADODB.Recordset
     
      With recSet
      .Open strSQL, Conn, 3, 2
      
        Do Until .EOF
        
         Set objList3 = ListView1.ListItems.Add(, , "Occupied")
             objList3.SubItems(1) = !booking_date
             objList3.SubItems(2) = !out_date
             objList3.SubItems(3) = !room_num
             objList3.SubItems(4) = !customer_fname & " " & !customer_lname
             
       .MoveNext
       
       Loop
      
      .Close
      
      End With
            
'for future booking
'------------------------------------------------------------------------------------------------------

strSQL = "SELECT tblBooking_Info.booking_date, tblBooking_Info.out_date, tblBooking_Info.room_num,  " & _
         "tblBooking_Info.booking_status, tblCustomer_Info.customer_fname,  " & _
         "tblCustomer_Info.customer_lname " & _
         "FROM tblRoom_Status INNER JOIN ((tblCustomer_Info INNER JOIN tblBooking_Record ON " & _
         "tblCustomer_Info.customer_num = tblBooking_Record.customer_num) INNER JOIN " & _
         "tblBooking_Info ON tblBooking_Record.booking_num = tblBooking_Info.booking_num) ON " & _
         "tblRoom_Status.room_num = tblBooking_Info.room_num " & _
         "WHERE (((tblBooking_Info.booking_date)=Date()) AND " & _
         "((tblBooking_Info.booking_status)=1)) OR " & _
         "(((tblBooking_Info.booking_date)>Date()));"


         
     Set recSet = New ADODB.Recordset
     
      With recSet
      
       .Open strSQL, Conn, 3, 2
       
        Do Until .EOF
        
         Set objList1 = ListView1.ListItems.Add(, , "Booking")
             objList1.SubItems(1) = !booking_date
             objList1.SubItems(2) = !out_date
             objList1.SubItems(3) = !room_num
             objList1.SubItems(4) = !customer_fname & " " & !customer_lname
             
       .MoveNext
       
       Loop
       
       .Close
      
      End With
      
'for future reservation
'-------------------------------------------------------------------------------------------------------

strSQL = "SELECT tblReservation_Info.reservation_date, tblReservation_Info.out_date, " & _
         "tblReservation_Info.room_num, tblReservation_Info.reservation_status,  " & _
         "tblCustomer_Info.customer_fname, tblCustomer_Info.customer_lname " & _
         "FROM tblCustomer_Info INNER JOIN (tblRoom_Status INNER JOIN " & _
         "(tblReservation_Record INNER JOIN tblReservation_Info ON  " & _
         "tblReservation_Record.reservation_num = tblReservation_Info.reservation_num) ON " & _
         "tblRoom_Status.room_num = tblReservation_Info.room_num) ON " & _
         "tblCustomer_Info.customer_num = tblReservation_Record.customer_num " & _
         "WHERE (((tblReservation_Info.reservation_date)=Date()) AND " & _
         "((tblReservation_Info.reservation_status)=1)) OR " & _
         "(((tblReservation_Info.reservation_date)>Date()));"

         
    Set recSet = New ADODB.Recordset
    
     With recSet
     
      .Open strSQL, Conn, 3, 2
      
         Do Until .EOF
        
         Set objList2 = ListView1.ListItems.Add(, , "Reservation")
             objList2.SubItems(1) = !reservation_date
             objList2.SubItems(2) = !out_date
             objList2.SubItems(3) = !room_num
             objList2.SubItems(4) = !customer_fname & " " & !customer_lname
             
       .MoveNext
       
       Loop
      
      .Close
          
    End With

End Sub

Private Sub txtGNumber_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtGNumber)
End Sub



Private Sub txtHeadCount_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtHeadCount)
End Sub
