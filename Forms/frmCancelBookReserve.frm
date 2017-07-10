VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmCancelBookReserve 
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   12375
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "frmCancelBookReserve.frx":0000
   ScaleHeight     =   5895
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcboIndex 
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin prjXTab.XTab XTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9551
      TabCaption(0)   =   "Tab 0"
      TabContCtrlCnt(0)=   7
      Tab(0)ContCtrlCap(1)=   "XPButton1"
      Tab(0)ContCtrlCap(2)=   "txtParameter"
      Tab(0)ContCtrlCap(3)=   "cboSearch"
      Tab(0)ContCtrlCap(4)=   "cmdViewAll"
      Tab(0)ContCtrlCap(5)=   "ListView1"
      Tab(0)ContCtrlCap(6)=   "Label2"
      Tab(0)ContCtrlCap(7)=   "Label1"
      TabCaption(1)   =   "Tab 1"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "ListView2"
      TabCaption(2)   =   "Tab 2"
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
      Begin XPControls.XPButton XPButton1 
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Search"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtParameter 
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   4320
         Width           =   1695
      End
      Begin XPControls.XPCombo cboSearch 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   3840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Text            =   "XPCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton cmdViewAll 
         Height          =   495
         Left            =   5520
         TabIndex        =   3
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&View All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   600
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   1
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
      Begin VB.Label Label2 
         Caption         =   "Search Parameter:"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Search By:"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   3720
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList imglstIcons1 
      Left            =   8400
      Top             =   4560
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
            Picture         =   "frmCancelBookReserve.frx":455B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":4DAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":2E9D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":585F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":58B8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":5E7AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":883D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":8F8D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":BA5AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":BAE87
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":BB761
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":BC03B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":BC1C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCancelBookReserve.frx":E5DE9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop1 
      Caption         =   "forReservation"
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Reservation"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReservationInfo 
         Caption         =   "Reservation Info"
      End
   End
   Begin VB.Menu mnuPop2 
      Caption         =   "forBoooking"
      Begin VB.Menu mnuCancelBooking 
         Caption         =   "Cancel Booking"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookInfo 
         Caption         =   "Booking Info"
      End
   End
End
Attribute VB_Name = "frmCancelBookReserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboSearch_Change()
txtcboIndex.Text = Trim$(cboSearch.listIndex)
End Sub

Private Sub cmdViewAll_Click()

Dim objList1 As ListItem
Dim objList2 As ListItem

ListView1.ListItems.Clear
ListView2.ListItems.Clear



strSQL = "SELECT tblBooking_Info.booking_date, tblBooking_Info.out_date, " & _
         "tblBooking_Info.room_num, tblBooking_Info.booking_status, " & _
         "tblCustomer_Info.customer_fname, tblCustomer_Info.customer_lname, " & _
         "tblBooking_Info.record_num, tblBooking_Info.expected_time_in, " & _
         "tblBooking_Info.expected_time_out, tblBilling_Booking.advance_payment, " & _
         "tblBilling_Booking.remaining_balance, tblBilling_Booking.total, " & _
         "tblBilling_Booking.refund, tblCustomer_Info.customer_num, tblBooking_Info.user_id, " & _
         "tblBooking_Info.check_in_time, tblBooking_Info.tran_type " & _
         "FROM tblRoom_Status INNER JOIN ((tblCustomer_Info INNER JOIN tblBooking_Record ON " & _
         "tblCustomer_Info.customer_num = tblBooking_Record.customer_num) INNER JOIN " & _
         "(tblBooking_Info INNER JOIN tblBilling_Booking ON " & _
         "tblBooking_Info.record_num = tblBilling_Booking.record_num) ON " & _
         "tblBooking_Record.booking_num = tblBooking_Info.booking_num) ON " & _
         "tblRoom_Status.room_num = tblBooking_Info.room_num " & _
         "WHERE (((tblBooking_Info.booking_date)=Date()) AND " & _
         "((tblBooking_Info.booking_status)=2)) OR (((tblBooking_Info.booking_date)>Date()));"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

If Not .EOF Then

    Do While Not .EOF
    
    i = 1
    
    Set objList1 = ListView1.ListItems.Add(i, , !tran_type & " Customer Name: " & !customer_fname & " " & !customer_lname, 2)
        objList1.SubItems(1) = !booking_date
        objList1.SubItems(2) = !out_date
        objList1.SubItems(3) = !room_num
        objList1.SubItems(4) = !customer_fname & " " & !customer_lname
        objList1.SubItems(5) = !record_num
        objList1.SubItems(6) = !advance_payment
        objList1.SubItems(7) = !remaining_balance
        objList1.SubItems(8) = !total
        objList1.SubItems(9) = !refund
        objList1.SubItems(10) = !customer_num
        objList1.SubItems(11) = !user_id
        objList1.SubItems(12) = !tran_type
        
        
        
    i = i + 1
    
    .MoveNext
    
    Loop

Else

MsgBox "No Records Found.", vbOKOnly, "No Records"

End If

End With
End Sub

Private Sub Form_Load()
FormPos frmCancelBookReserve
DBCON
TabProp
List_ViewProp
End Sub

Private Sub TabProp()

With XTab1
.TabCount = 2
.TabCaption(0) = "Booking"
.TabCaption(1) = "Reservation"

End With

With cboSearch
.Text = "Booking Number"
.AddItem "Booking Number", 0
.AddItem "Customer Name", 1
.AddItem "Booking Date", 2

End With
End Sub

Private Sub List_ViewProp()

Set ListView1.Icons = imglstIcons1
Set ListView2.Icons = imglstIcons1

With ListView1

.Arrange = lvwAutoTop
.HotTracking = True
.LabelEdit = lvwManual
.BorderStyle = ccFixedSingle
.Appearance = ccFlat
.OLEDragMode = ccOLEDragAutomatic

For a = 1 To 15
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

For c = 1 To 15
.ColumnHeaders.Add c
Next c

End With
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then

PopupMenu mnuPop2

End If

End Sub

Private Sub mnuBookInfo_Click()
On Error GoTo ErrCatch




ErrCatch:
Select Case Err.Number

Case 91

MsgBox "No Records Found. Please Try Again.", vbCritical + vbOKOnly, "Error"

Exit Sub

End Select
End Sub

Private Sub mnuCancelBooking_Click()
On Error GoTo ErrCatch

ErrCatch:
Select Case Err.Number

Case 91

MsgBox "No Records Found. Please Try Again.", vbCritical + vbOKOnly, "Error"

Exit Sub

End Select
End Sub

Private Sub XPButton1_Click()

If txtcboIndex.Text = "0" Or txtcboIndex.Text = "-1" Then

MsgBox "Booking number"

ElseIf txtcboIndex.Text = 1 Then

MsgBox "Customer name"

ElseIf txtcboIndex.Text = 2 Then

MsgBox "booking date"

End If

End Sub
