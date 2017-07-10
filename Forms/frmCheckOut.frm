VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmCheckOut 
   Caption         =   "Check-Out"
   ClientHeight    =   4515
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7620
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "frmCheckOut.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin XPControls.XPButton cmdBack 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   3840
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6165
      Arrange         =   2
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
   Begin MSComctlLib.ImageList imglstIcons1 
      Left            =   9360
      Top             =   4800
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
            Picture         =   "frmCheckOut.frx":455B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":4DAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":2E9D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":585F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":58B8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":5E7AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":883D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":8F8D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":BA5AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":BAE87
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":BB761
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":BC03B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":BC1C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckOut.frx":E5DE9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckOut 
         Caption         =   "Check-out"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckInAndOut 
         Caption         =   "InAndOut"
      End
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ttype As String
Dim RoomNumber As Integer
Dim CustomerNumber As Integer
Dim RecordNumber As Integer
Dim CheckInDate, CheckOutDate As Date


Private Sub Form_Load()
FormPos frmCheckOut
DBCON
ListViewProp
End Sub

Private Sub ListViewProp()

Dim objList As ListItem
Set ListView1.Icons = imglstIcons1

With ListView1
.Arrange = lvwAutoTop
.HotTracking = True
.HoverSelection = True
.LabelEdit = lvwManual
.BorderStyle = ccFixedSingle
.Appearance = ccFlat
.OLEDragMode = ccOLEDragAutomatic

For a = 1 To 7
.ColumnHeaders.Add a
Next a
End With

strSQL = "SELECT tblCustomer_Info.customer_fname, tblCustomer_Info.customer_lname, " & _
         "tblRoom_Status.room_num, tblRoom_Status.customer_number, tblBooking_Info.booking_date, " & _
         "tblBooking_Info.out_date, tblBooking_Info.check_in_time, " & _
         "tblBooking_Info.expected_time_out, tblRoom_Status.room_status, " & _
         "tblBooking_Info.tran_type, tblBooking_Info.booking_status, " & _
         "tblBooking_Info.record_num, tblBooking_Info.tran_type " & _
         "FROM (tblCustomer_Info INNER JOIN tblRoom_Status ON " & _
         "tblCustomer_Info.customer_num = tblRoom_Status.customer_number) INNER JOIN " & _
         "tblBooking_Info ON tblRoom_Status.room_num = tblBooking_Info.room_num " & _
         "WHERE (((tblRoom_Status.room_status)=1) AND ((tblBooking_Info.booking_status)=0));"




         
 Set recSet = New ADODB.Recordset
 
    With recSet
    .Open strSQL, Conn, 3, 3
    
    Do While Not .EOF
    i = 1
    
 Set objList = ListView1.ListItems.Add(i, , "Room " & !room_num & " " & !customer_fname & " " & !customer_lname, 8)
     objList.SubItems(1) = !room_num 'room number
     objList.SubItems(2) = !customer_number 'customer number
     objList.SubItems(3) = !tran_type
     objList.SubItems(4) = !record_num
     objList.SubItems(5) = "Check-In Date: " & !booking_date & " Check-Out Date: " & !out_Date & _
                           " Expected Time Out: " & !expected_time_out
     
     
    i = i + 1
    
    .MoveNext
    Loop
    .Close
    
    
    End With
         


End Sub


Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then

frmCheckOut.mnuCheckInAndOut.Caption = ListView1.SelectedItem.SubItems(5)

PopupMenu frmCheckOut.mnuPopUp

End If
End Sub

Private Sub mnuCheckOut_Click()
If MsgBox("Are you sure you want to check out this guest? This process can't be reverted back.", vbInformation + vbYesNo, "Confirm Check out") = vbYes Then

RoomNumber = ListView1.SelectedItem.SubItems(1)
CustomerNumber = ListView1.SelectedItem.SubItems(2)
Ttype = ListView1.SelectedItem.SubItems(3)
RecordNumber = ListView1.SelectedItem.SubItems(4)

strSQL = "select * from tbldatesandtime where transaction_type='" & Ttype & "'"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

If Not .EOF Then

.Delete
.Update

End If

End With

strSQL = "SELECT tblRoom_Status.room_num, tblRoom_Status.room_status, " & _
         "tblRoom_Status.customer_number, * " & _
         "From tblRoom_Status " & _
         "WHERE (((tblRoom_Status.room_num)=" & RoomNumber & ") AND " & _
         "((tblRoom_Status.room_status)=1) AND ((tblRoom_Status.customer_number)=" & CustomerNumber & "));"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

If Not .EOF Then

!room_status = 0
!customer_number = Null

.Update

End If

End With

strSQL = "select * from tblbooking_info where record_num=" & RecordNumber & ""

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

If Not .EOF Then

!check_out_time = Format$(Now, "hh:mm AM/PM")
!booking_status = 1

.Update

.Close

End If
End With

Else

Exit Sub

End If

MsgBox "Guest has been successfully checked-out", vbInformation + vbOKOnly, "Check-out Success"

ListView1.ListItems.Clear

ListViewProp

End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub
