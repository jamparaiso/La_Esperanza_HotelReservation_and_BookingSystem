VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{E0D30636-0F87-47D5-B501-08A4FFAC604E}#1.0#0"; "osenxpsuite2005.OCX"
Begin VB.Form frmCheckIn 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8610
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCheckIn.frx":0000
   ScaleHeight     =   8610
   ScaleWidth      =   11070
   Begin VB.TextBox txtFanRoomPax 
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
      Left            =   2040
      TabIndex        =   46
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtDateCheckout 
      Height          =   495
      Left            =   7560
      TabIndex        =   44
      Top             =   10560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCmdConfirmTag 
      Height          =   495
      Left            =   11520
      TabIndex        =   43
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin osenxpsuite2005.OsenXPDTPicker dtpickCheckIn 
      Height          =   315
      Left            =   2040
      TabIndex        =   42
      Top             =   5280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Text            =   "2012-09-20"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      YEAR            =   0
      MONTH           =   0
      MYDATE          =   0
      thisdate        =   41172
   End
   Begin VB.TextBox txtCheckOutTimeIndex 
      Height          =   495
      Left            =   11520
      TabIndex        =   39
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRoomRate 
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
      Left            =   2040
      TabIndex        =   38
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtCurrentMonth 
      Height          =   495
      Left            =   7560
      TabIndex        =   36
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDateDiff 
      Height          =   495
      Left            =   7560
      TabIndex        =   35
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   8880
      TabIndex        =   34
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   64159745
      CurrentDate     =   41165
   End
   Begin VB.TextBox txtRoomNum 
      Height          =   495
      Left            =   8880
      TabIndex        =   33
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTimeIn 
      Height          =   495
      Left            =   10200
      TabIndex        =   32
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin XPControls.XPCombo xpcboExtendTime 
      Height          =   315
      Left            =   2040
      TabIndex        =   31
      Top             =   6480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Text            =   "Choose Here"
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
   Begin VB.TextBox txtTotalAmenities 
      Height          =   495
      Left            =   11520
      TabIndex        =   30
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCurrentDate 
      Height          =   495
      Left            =   8880
      TabIndex        =   29
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRoomTypeIndex 
      Height          =   495
      Left            =   10200
      TabIndex        =   28
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin XPControls.XPButton xpbtRemove 
      Height          =   495
      Left            =   6720
      TabIndex        =   27
      Top             =   6240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "&Remove Selected Item"
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
   Begin VB.TextBox txtQuantity 
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
      Left            =   9360
      TabIndex        =   26
      Top             =   1680
      Width           =   1215
   End
   Begin XPControls.XPButton xpbtAdd 
      Height          =   495
      Left            =   8640
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Add"
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
   Begin VB.TextBox txtItemPrice 
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
      Left            =   6480
      TabIndex        =   23
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtItemName 
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
      Left            =   6480
      TabIndex        =   22
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtItemNum 
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
      Left            =   6480
      TabIndex        =   21
      Top             =   1680
      Width           =   1455
   End
   Begin XPControls.XPButton xpbtClear 
      Height          =   495
      Left            =   4560
      TabIndex        =   16
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Clear"
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
   Begin XPControls.XPButton xpbtConfirm 
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Confirm"
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
   Begin osenxpsuite2005.OsenXPDTPicker dtpickCheckOut 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   5880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Text            =   "2012-09-11"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      YEAR            =   0
      MONTH           =   0
      MYDATE          =   0
      thisdate        =   41163
   End
   Begin XPControls.XPCombo xpcboRoomNumber 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Top             =   4680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Text            =   "XPCombo2"
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
   Begin XPControls.XPCombo xpcboRoomType 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Text            =   "XPCombo1"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtGnumber 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   4800
      TabIndex        =   8
      Top             =   3840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3836
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin XPControls.XPButton xpbtBack 
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label lblTranType 
      BackStyle       =   0  'Transparent
      Caption         =   "House Registration Section"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   48
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label lblnumguest 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of guest:"
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
      TabIndex        =   47
      Top             =   4080
      Width           =   1920
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   6735
      Left            =   4440
      Top             =   960
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   6735
      Left            =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   45
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmCheckIn.frx":455B
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11055
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-In Date:"
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
      TabIndex        =   41
      Top             =   5280
      Width           =   1680
   End
   Begin VB.Label lblroomrate 
      BackStyle       =   0  'Transparent
      Caption         =   "* room rate vary on peak season"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   40
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Rate:"
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
      TabIndex        =   37
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
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
      Left            =   8160
      TabIndex        =   25
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Price"
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
      Left            =   5040
      TabIndex        =   20
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Number:"
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
      Left            =   5040
      TabIndex        =   18
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chosen Amenities:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   4680
      TabIndex        =   14
      Top             =   3480
      Width           =   2550
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amenities:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   4680
      TabIndex        =   7
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-out time:"
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
      Top             =   6480
      Width           =   1800
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Number:"
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
      TabIndex        =   5
      Top             =   4680
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Type:"
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
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-Out Date:"
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
      TabIndex        =   3
      Top             =   5880
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guest Name:"
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
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guest Number:"
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
      TabIndex        =   1
      Top             =   1800
      Width           =   1560
   End
   Begin VB.Label lblControl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "checkin"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1050
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FanRate As Integer
Private Sub dtpickCheckIn_LostFocus()
Dim CurrentDateNow As Date
CurrentDateNow = Format$(Now, "mm/dd/yyyy")
If dtpickCheckIn.Value < CurrentDateNow Then
MsgBox "You Cant Choose Date Later Than Today", vbCritical + vbOKOnly, "Invalid Check-Out Date"
dtpickCheckIn.Value = Format$(Now, "mm/dd/yyyy")
dtpickCheckIn.SetFocus
Exit Sub

Else
DTPicker1.Value = dtpickCheckIn.Value
DTPicker1_Change
txtDateDiff.Text = dtpickCheckOut.Value - DTPicker1.Value
Exit Sub
End If
End Sub

Private Sub Form_Load()
FormPos frmCheckIn
DBCON
FormControlProperties
ListViewProp
RoomTypes
xpcboExtendItems


End Sub

Private Sub VerifyRooms()

strSQL = "select * from tbldatesandtime where room_number=" & xpcboRoomNumber.Text & ""

Set recSet = New ADODB.Recordset

    With recSet
    .Open strSQL, Conn, 3, 3
    
    Do While Not .EOF
        
     If DateValue(txtCurrentDate.Text) >= DateValue(!date_check_in) And _
        DateValue(txtCurrentDate.Text) <= DateValue(!date_check_out) And _
        TimeNo(xpcboExtendTime.Text) >= TimeNo(!check_in_time) Then
            'this is for when the guest will check-in an occupied room that will check-out on the same date.
            If txtCurrentDate.Text = !date_check_out And TimeNo(txtTimeIn.Text) > TimeNo(!check_out_time) Then

                If MsgBox("Room is Available. Before Continuing please make sure that all the " & _
                          "information supplied is true and valid. Do you want to continue?", _
                           vbInformation + vbYesNo, "Confirm Information") = vbYes Then
                
                 frmBillingCheckIn.Show vbModal, frmCheckIn
                 
                 Exit Sub

                 
                Else
                
                    Exit Sub
                
                End If
            
            Else
            'if the guest currently occupying the room has extended his checkout time.
            'thus conflicting the check-in time of the oncoming guest and the checkout time of the current occupant
            MsgBox "This Room Will Be Unavailable from " & !check_in_time & " of " _
                    & !date_check_in & " until " & !check_out_time & " of " _
                    & !date_check_out & ".RF Number" & !transaction_type, vbCritical + vbOKOnly, "Reserved Room"
            
              Exit Sub
            
            End If
        
        
     ElseIf DateValue(txtCurrentDate.Text) >= DateValue(!date_check_in) And _
        DateValue(txtCurrentDate.Text) <= DateValue(!date_check_out) And _
        TimeNo(xpcboExtendTime.Text) < TimeNo(!check_in_time) Then
        
            If DateValue(txtCurrentDate.Text) = DateValue(!date_check_out) And TimeNo(txtTimeIn.Text) > TimeNo(!check_out_time) Then
            
                If MsgBox("Room is Available. Before Continuing please make sure that all the " & _
                          "information supplied is true and valid. Do you want to continue?", _
                           vbInformation + vbYesNo, "Confirm Information") = vbYes Then
                
                frmBillingCheckIn.Show vbModal, frmCheckIn
                
                Exit Sub
                 
                Else
                
                    Exit Sub
                
                End If
            
            Else
            
            MsgBox "This Room Will Be Unavailable from " & !check_in_time & " of " _
                    & !date_check_in & " until " & !check_out_time & " of " _
                    & !date_check_out & ".RF Number" & !transaction_type, vbCritical + vbOKOnly, "Reserved Room"
                    
                Exit Sub
            
            End If
            
     ElseIf DateValue(dtpickCheckOut.Value) >= DateValue(!date_check_in) And _
            DateValue(dtpickCheckOut.Value) <= DateValue(!date_check_out) And _
            TimeNo(xpcboExtendTime.Text) >= TimeNo(!check_in_time) Then
            
            MsgBox "This Room Will Be Unavailable from " & !check_in_time & " of " _
                    & !date_check_in & " until " & !check_out_time & " of " _
                    & !date_check_out & ".RF Number" & !transaction_type, vbCritical + vbOKOnly, "Reserved Room"
                    
              Exit Sub
           
     ElseIf DateValue(dtpickCheckOut.Value) >= DateValue(!date_check_in) And _
            DateValue(dtpickCheckOut.Value) <= DateValue(!date_check_out) And _
            TimeNo(xpcboExtendTime.Text) < TimeNo(!check_in_time) Then
            
            If DateValue(dtpickCheckOut.Value) = DateValue(!date_check_in) Then
            
                If MsgBox("Room is Available. Before Continuing please make sure that all the " & _
                          "information supplied is true and valid. Do you want to continue?", _
                           vbInformation + vbYesNo, "Confirm Information") = vbYes Then
                
                 frmBillingCheckIn.Show vbModal, frmCheckIn
                 
                 Exit Sub
                 
                Else
                
                    Exit Sub
                
                End If
            
            Else
            
            MsgBox "This Room Will Be Unavailable from " & !check_in_time & " of " _
                    & !date_check_in & " until " & !check_out_time & " of " _
                    & !date_check_out & ".RF Number" & !transaction_type, vbCritical + vbOKOnly, "Reserved Room"
                    
                    Exit Sub
            
            End If
            
     ElseIf DateValue(txtCurrentDate.Text) <= !date_check_in And _
            DateValue(dtpickCheckOut.Value) >= !date_check_out Then
            
            MsgBox "This Room Will Be Unavailable from " & !check_in_time & " of " _
                    & !date_check_in & " until " & !check_out_time & " of " _
                    & !date_check_out & ".RF Number" & !transaction_type, vbCritical + vbOKOnly, "Reserved Room"

                    
              Exit Sub
                 
     End If
     .MoveNext
     Loop
    
    .Close
    End With
    
    If MsgBox("Room is Available. Before Continuing please make sure that all the " & _
                          "information supplied is true and valid. Do you want to continue?", _
                           vbInformation + vbYesNo, "Confirm Information") = vbYes Then
                
                 frmBillingCheckIn.Show vbModal, frmCheckIn
                 
                 Exit Sub
                 
    Else
                
        Exit Sub
        
    End If
    
End Sub


Private Sub RoomTypes()

'get all the list of rooms in db and assign a index that to be used in cboroomnumbeer
'---------------------------------------------------------------------------------------------------
strSQL = "select * from tblroom_info"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 2
  
   i = 0
  
   Do Until .EOF
   
    xpcboRoomType.AddItem !room_name, i 'i is the index
    
    .MoveNext
    
    i = i + 1
   
   Loop
  
  .Close
  
  End With
 
End Sub

Private Sub RoomList()
Dim roomIndex As Integer

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
If txtCmdConfirmTag.Text = "1" Then

strSQL = "select * from tblroom_status where room_type=" & roomIndex & " and room_status= 0"


 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   Do Until .EOF
   
   xpcboRoomNumber.AddItem !room_num
   
   .MoveNext
   
   Loop
  
  .Close
  
  End With
  
Else

strSQL = "select * from tblroom_status where room_type=" & roomIndex & ""


 Set recSet = New ADODB.Recordset
 
  With recSet
  
  .Open strSQL, Conn, 3, 3
  
   Do Until .EOF
   
   xpcboRoomNumber.AddItem !room_num
   
   .MoveNext
   
   Loop
  
  .Close
  
  End With
  
End If
  
'for rates, pax, etc,,,
'-----------------------------------------------------------------------------------------------------

 If txtCurrentMonth.Text = "1" Then 'if the month is january(prior to ati-atihan)the room rates will change
 
  strSQL = "select * from tblroom_info where room_type_num= " & roomIndex & " "
  
   Set recSet = New ADODB.Recordset
   
    With recSet
     .Open strSQL, Conn, 3, 2
     
      txtRoomRate.Text = !room_rate_on_season
     
     .Close
     
    End With


 Else 'if not january then normal rate are applied

    strSQL = "select * from tblroom_info where room_type_num= " & roomIndex & " "

        Set recSet = New ADODB.Recordset
 
            With recSet
  
                .Open strSQL, Conn, 3, 2
  
                    txtRoomRate.Text = !room_rate
  
                .Close
  
            End With

 End If
End Sub
Private Sub xpcboExtendItems()

With xpcboExtendTime
.AddItem "12:00 PM", 0
.AddItem "1:00 PM", 1
.AddItem "2:00 PM", 2
.AddItem "3:00 PM", 3
.AddItem "4:00 PM", 4
.AddItem "5:00 PM", 5
.AddItem "6:00 PM", 6
.Text = "12:00 PM"
End With

End Sub

Private Sub SearchItem()

strSQL = "select * from tblamenities where item_num=" & txtItemNum.Text & ""

    Set recSet = New ADODB.Recordset
    
        With recSet
        .Open strSQL, Conn, 3, 2
        
        txtItemName.Text = !item_name
        txtItemPrice.Text = !item_price
        
        .Close
        
        End With

End Sub
Private Sub FormControlProperties()

txtGName.Locked = True
xpcboRoomType.AutoSearch = True
xpcboRoomNumber.AutoSearch = True
xpcboRoomType.Text = "Choose Here"
xpcboRoomNumber.Text = "Choose Here"
txtGName.Locked = True
dtpickCheckOut.FormatDate = "mm/dd/yyyy"
dtpickCheckIn.FormatDate = "mm/dd/yyyy"
dtpickCheckOut.Value = Format(Now + 1, "mm/dd/yyyy")
dtpickCheckIn.Value = Format(Now, "mm/dd/yyyy")
ListView1.View = lvwReport
txtItemName.Locked = True
txtItemPrice.Locked = True
txtQuantity.Text = 0
xpbtRemove.Enabled = False
txtRoomTypeIndex.Visible = True
txtCurrentDate.Text = Format(Now, "mm/dd/yyyy")
txtTimeIn.Text = Format("2:00 PM", "hh:mm AM/PM")
DTPicker1.Format = dtpShortDate
DTPicker1.Value = Format(Now, "mm/dd/yyyy")
txtCurrentDate.Text = DTPicker1.Value
txtDateDiff.Text = dtpickCheckOut.Value - DTPicker1.Value
txtCurrentMonth.Text = Month(Now)
txtTotalAmenities.Text = 0
txtRoomRate.Locked = True
lblnumguest.Visible = False
txtFanRoomPax.Visible = False
lblroomrate.Visible = True

End Sub

Private Sub ListViewProp()

    With ListView1
    
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Item Num", .Width * 0.15
        .ColumnHeaders.Add 2, , "Item Name", .Width * 0.25
        .ColumnHeaders.Add 3, , "Item Price", .Width * 0.2
        .ColumnHeaders.Add 4, , "Quantity", .Width * 0.15
        .ColumnHeaders.Add 5, , "Total", .Width * 0.24
        
    End With
End Sub

Private Sub ViewListAddItem()
Dim objList As ListItem

Set objList = ListView1.ListItems.Add(, , txtItemNum.Text)
    objList.SubItems(1) = txtItemName.Text
    objList.SubItems(2) = txtItemPrice.Text
    objList.SubItems(3) = txtQuantity.Text
    objList.SubItems(4) = Val(txtItemPrice.Text) * Val(txtQuantity.Text)


End Sub



Private Sub ListView1_Click()
xpbtRemove.Enabled = True
End Sub

Private Sub SearchGuest()
If txtGnumber.Text <> "" Then

strSQL = "select * from tblcustomer_info where customer_num = " & txtGnumber.Text & ""

 Set recSet = New ADODB.Recordset
 
    With recSet
    .Open strSQL, Conn, 3, 2
    If .EOF = False Then
    txtGName.Text = !customer_fname & " " & !customer_lname
    Exit Sub
    Else
    txtGName.Text = ""
    
    End If
    .Close
    
    End With
Else

txtGName.Text = ""

End If
End Sub



Private Sub txtFanRoomPax_Change()
txtRoomRate.Text = Val(FanRate) * Val(txtFanRoomPax)
End Sub

Private Sub txtFanRoomPax_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtFanRoomPax)
End Sub

Private Sub txtFanRoomPax_LostFocus()
If txtFanRoomPax.Text = "0" Or txtFanRoomPax.Text = "" Then
MsgBox "Please put the correct number of occupants.", vbCritical + vbOKOnly, "Error"
txtFanRoomPax.Text = 1
txtFanRoomPax.SetFocus
End If
End Sub

Private Sub txtGNumber_Change()
SearchGuest
txtCmdConfirmTag.Text = xpbtConfirm.Tag
End Sub

Private Sub txtGNumber_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtGnumber)
End Sub


Private Sub txtItemNum_Change()
If txtItemNum.Text <> "" Then
SearchItem
Else
txtItemName.Text = ""
txtItemPrice.Text = ""
End If
End Sub

Private Sub txtItemNum_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtItemNum)
End Sub

Private Sub txtQuantity_GotFocus()
txtQuantity.Text = ""
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtQuantity)
End Sub


Private Sub txtQuantity_LostFocus()
If txtQuantity.Text = "" Then
txtQuantity.Text = 0
End If
End Sub

Private Sub xpbtAdd_Click()
If txtItemName.Text <> "" Then
    If txtQuantity.Text <> 0 Or txtQuantity.Text = "" Then
    
        ViewListAddItem
        
        ColumnTotalListView
        
        txtItemNum.Text = ""
        txtQuantity.Text = 0
        
    Else
        MsgBox "Invalid value for quantity.Must be higher than zero.Try Again.", vbInformation + vbOKOnly, "Invalid Quantity"
        txtQuantity.SetFocus
    End If
Else
MsgBox "Invalid Item Number!Please Try Again", vbCritical + vbOKOnly, "No Records Found"
txtItemNum.Text = ""
txtItemNum.SetFocus
End If
End Sub

Private Sub ColumnTotalListView()
Dim listIndex As Long
    Dim listTotal As Long
    
    For listIndex = 1 To ListView1.ListItems.Count
        listTotal = listTotal + ListView1.ListItems(listIndex).SubItems(4)
    Next
    
    txtTotalAmenities.Text = listTotal

End Sub

Private Sub xpbtBack_Click()
Unload Me
End Sub

Private Sub xpbtConfirm_Click()
On Error GoTo Errorcatch

If txtGName.Text <> "" Then

VerifyRooms
If txtCmdConfirmTag.Text = "3" Then
frmBillingCheckIn.xpbtCheckIn.Enabled = True
End If

Else
MsgBox "Please Put The Correct Customer Number.", vbCritical + vbOKOnly, "No Customer Found"
txtGnumber.SetFocus
Exit Sub

End If

Errorcatch:
Select Case Err.Number

Case 13
MsgBox "Please Select the Check-Out Time of the Guest", vbCritical + vbOKOnly, "No Check-Out Time Selected"
Exit Sub

Case -2147217900
MsgBox "Please Select the Room Type or the Room Number", vbCritical + vbOKOnly, "No Room Type or Room Number Selected"
Exit Sub

End Select
End Sub

Private Sub xpbtRemove_Click()

  If ListView1.SelectedItem Is Nothing Then Exit Sub

    If MsgBox("Do you really want to delete this amenity?", vbInformation + vbYesNo, "Confirm Delete") = vbYes Then
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
        ColumnTotalListView
         xpbtRemove.Enabled = False
    End If

    
End Sub


Private Sub xpcboExtendTime_Change()
txtCheckOutTimeIndex.Text = xpcboExtendTime.listIndex
End Sub

Private Sub xpcboRoomNumber_Change()
txtRoomNum.Text = xpcboRoomNumber.Text
End Sub

Private Sub xpcboRoomType_Change()
txtRoomTypeIndex.Text = xpcboRoomType.listIndex
xpcboRoomNumber.Clear
txtCmdConfirmTag.Text = xpbtConfirm.Tag
RoomList
If txtRoomTypeIndex.Text = 0 Then
lblnumguest.Visible = True
txtFanRoomPax.Visible = True
lblroomrate.Visible = False
FanRate = txtRoomRate.Text
Else
lblnumguest.Visible = False
txtFanRoomPax.Visible = False
lblroomrate.Visible = True
End If
End Sub

Private Sub DTPicker1_Change()
txtCurrentDate.Text = DTPicker1.Value
End Sub



Private Sub dtpickCheckOut_LostFocus()

If dtpickCheckOut.Value < DTPicker1.Value Then
MsgBox "You Cant Choose Date Later Than Your Check-in date", vbCritical + vbOKOnly, "Invalid Check-Out Date"
dtpickCheckOut.Value = DTPicker1.Value + 1
dtpickCheckOut.SetFocus
Exit Sub

Else
txtDateDiff.Text = dtpickCheckOut.Value - DTPicker1.Value
'txtDateCheckout.Text = dtpickCheckOut.Value
Exit Sub
End If
End Sub

