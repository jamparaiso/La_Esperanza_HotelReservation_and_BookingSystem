VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{E0D30636-0F87-47D5-B501-08A4FFAC604E}#1.0#0"; "osenxpsuite2005.OCX"
Begin VB.Form frmBooking 
   Caption         =   "Booking"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form2"
   ScaleHeight     =   6960
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGnumber 
      Height          =   495
      Left            =   1440
      TabIndex        =   23
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtGName 
      Height          =   495
      Left            =   4320
      TabIndex        =   22
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtItemNum 
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtItemName 
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtItemPrice 
      Height          =   495
      Left            =   7440
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtQuantity 
      Height          =   495
      Left            =   9960
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtRoomTypeIndex 
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtCurrentDate 
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtTotalAmenities 
      Height          =   495
      Left            =   9960
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtTimeIn 
      Height          =   495
      Left            =   8640
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtRoomNum 
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtDateDiff 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtCurrentMonth 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtRoomRate 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtCheckOutTimeIndex 
      Height          =   495
      Left            =   9960
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   7320
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   41165
   End
   Begin XPControls.XPCombo xpcboExtendTime 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Text            =   "Choose Here"
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
   Begin XPControls.XPButton xpbtRemove 
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "&Remove Selected Item"
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
   Begin XPControls.XPButton xpbtAdd 
      Height          =   495
      Left            =   8880
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Add"
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
   Begin XPControls.XPButton xpbtClear 
      Height          =   495
      Left            =   1800
      TabIndex        =   17
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Clear"
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
   Begin XPControls.XPButton xpbtConfirm 
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Confirm"
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
   Begin osenxpsuite2005.OsenXPDTPicker dtpickCheckOut 
      Height          =   315
      Left            =   1440
      TabIndex        =   19
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Text            =   "2012-09-11"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Left            =   1440
      TabIndex        =   20
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Text            =   "XPCombo2"
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
   Begin XPControls.XPCombo xpcboRoomType 
      Height          =   315
      Left            =   1440
      TabIndex        =   21
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   1335
      Left            =   6480
      TabIndex        =   24
      Top             =   3120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2355
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
      Left            =   3240
      TabIndex        =   25
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Back"
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
   Begin VB.Label Label1 
      Caption         =   "Check-In"
      Height          =   195
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Guest Number:"
      Height          =   495
      Left            =   0
      TabIndex        =   39
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Guest Name:"
      Height          =   495
      Left            =   2880
      TabIndex        =   38
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Check-Out Date:"
      Height          =   495
      Left            =   0
      TabIndex        =   37
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Room Type:"
      Height          =   495
      Left            =   0
      TabIndex        =   36
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Room Number:"
      Height          =   495
      Left            =   0
      TabIndex        =   35
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Check-out time:"
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Amenities:"
      Height          =   495
      Left            =   7200
      TabIndex        =   33
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Chosen Amenities:"
      Height          =   495
      Left            =   5400
      TabIndex        =   32
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Item Number:"
      Height          =   495
      Left            =   6000
      TabIndex        =   31
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Item Name:"
      Height          =   495
      Left            =   6000
      TabIndex        =   30
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Item Price"
      Height          =   495
      Left            =   6000
      TabIndex        =   29
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Quantity:"
      Height          =   495
      Left            =   8760
      TabIndex        =   28
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "*Room Rate:"
      Height          =   495
      Left            =   0
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "* room rate vary on peak season"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   2535
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
