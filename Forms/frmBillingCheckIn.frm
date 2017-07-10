VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmBillingCheckIn 
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmBillingCheckIn.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin XPControls.XPButton cmdPrint 
      Height          =   495
      Left            =   5520
      TabIndex        =   54
      Top             =   7440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Print OR"
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
   Begin XPControls.XPFrame XPFrame3 
      Height          =   2655
      Left            =   5280
      TabIndex        =   47
      Top             =   4080
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4683
      BackColor       =   8454016
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtReservationTotal 
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
         Left            =   1440
         TabIndex        =   49
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBillingCheckIn.frx":455B
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         TabIndex        =   50
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:*"
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
         TabIndex        =   48
         Top             =   120
         Width           =   840
      End
   End
   Begin XPControls.XPFrame XPFrame2 
      Height          =   3135
      Left            =   5280
      TabIndex        =   38
      Top             =   4080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      BackColor       =   8454016
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox chkboxBookGreaterPay 
         BackColor       =   &H0080FF80&
         Caption         =   "check this if the customer wants to pay the full amount or greater than the amount of ""Amount To Pay"""
         Height          =   1455
         Left            =   3120
         TabIndex        =   53
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtBookNextPay 
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
         Left            =   1560
         TabIndex        =   51
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtBookChange 
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
         Left            =   1560
         TabIndex        =   46
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtBookPaid 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   44
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtBookPay 
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
         Left            =   1560
         TabIndex        =   42
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtBookTotal 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   40
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount To Pay on arrival date:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount To Pay:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.TextBox txtCheckOutTimePayment 
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
      Left            =   6960
      TabIndex        =   37
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtTotalAmenitiesPayment 
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
      Left            =   6960
      TabIndex        =   36
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtTotalRoomPayment 
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
      Left            =   6960
      TabIndex        =   35
      Top             =   1680
      Width           =   1215
   End
   Begin XPControls.XPFrame XPFrame1 
      Height          =   1815
      Left            =   5280
      TabIndex        =   24
      Top             =   4080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3201
      BackColor       =   8454016
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   1680
         TabIndex        =   27
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtAmoutPaid 
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
         Left            =   1680
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
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
         Left            =   1680
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label13 
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
         TabIndex        =   30
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label14 
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
         TabIndex        =   29
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label15 
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
         TabIndex        =   28
         Top             =   1320
         Width           =   840
      End
   End
   Begin XPControls.XPButton xpbtCheckIn 
      Height          =   495
      Left            =   7440
      TabIndex        =   22
      Top             =   7440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Check-In"
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
   Begin VB.TextBox txtCheckOutTime 
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
      TabIndex        =   21
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtCheckInTime 
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
      TabIndex        =   20
      Top             =   6000
      Width           =   1575
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
      TabIndex        =   19
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtCheckOutDate 
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
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtCheckInDate 
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
      TabIndex        =   17
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtRoomNumber 
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
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtRoomType 
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
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
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
      TabIndex        =   14
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtGNumber 
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
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtRFNumber 
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
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   720
      TabIndex        =   10
      Top             =   7560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2143
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
      Left            =   1320
      TabIndex        =   23
      Top             =   8880
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   2895
      Left            =   4560
      Top             =   960
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   8535
      Left            =   0
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Statement"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1560
      TabIndex        =   55
      Top             =   120
      Width           =   4890
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmBillingCheckIn.frx":4642
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-Out Time Extension:"
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
      Left            =   5640
      TabIndex        =   34
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amenities:"
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
      Left            =   5640
      TabIndex        =   33
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room:"
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
      Left            =   5640
      TabIndex        =   32
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Breakdown:"
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
      Left            =   4800
      TabIndex        =   31
      Top             =   1200
      Width           =   2700
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amenities Requested:"
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
      TabIndex        =   11
      Top             =   7200
      Width           =   2400
   End
   Begin VB.Label Label11 
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
      TabIndex        =   9
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-Out Time:"
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
      Top             =   6600
      Width           =   1800
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-In Time:"
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
      TabIndex        =   7
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Label Label8 
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
      TabIndex        =   6
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label Label7 
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
      TabIndex        =   5
      Top             =   4800
      Width           =   1680
   End
   Begin VB.Label Label6 
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
      TabIndex        =   4
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Label Label5 
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
      TabIndex        =   3
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Label Label4 
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
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   1800
      Width           =   1560
   End
End
Attribute VB_Name = "frmBillingCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BookingNumber As Integer
Dim ReservationNumber As Integer
Dim BookingMax As Integer
Dim ReservationMax As Integer
Dim BillingMax As Integer
Dim objList As ListItem
Dim ListCount As Integer
Dim PerHourRate As Integer

Private Sub chkboxBookGreaterPay_Click()
If chkboxBookGreaterPay.Value = 1 Then
txtBookPay.Locked = False
MsgBox "Please put the amount in 'Amount To Pay'", vbOKOnly, "Input Amount"
txtBookPay.Text = ""
txtBookPay.SetFocus
End If
End Sub

Private Sub Form_Load()
FormPos frmBillingCheckIn
DBCON

FormLoadItems
GetCustomerName
ListViewProp
MigrateListItem
GetRFNumFromBookingInfo
CheckcmdConfirmTag

End Sub

Private Sub CheckcmdConfirmTag()

If frmCheckIn.txtCmdConfirmTag.Text = "1" Then
xpbtCheckIn.Caption = "&Check-In"
XPFrame1.Visible = True
XPFrame2.Visible = False
XPFrame3.Visible = False

ElseIf frmCheckIn.txtCmdConfirmTag.Text = "2" Then
xpbtCheckIn.Caption = "&Book"
XPFrame1.Visible = False
XPFrame2.Visible = True
XPFrame3.Visible = False


ElseIf frmCheckIn.txtCmdConfirmTag.Text = "3" Then
xpbtCheckIn.Caption = "&Reserve"
XPFrame1.Visible = False
XPFrame2.Visible = False
XPFrame3.Visible = True
xpbtCheckIn.Enabled = True
txtCheckInTime.Text = "2:00 PM"


End If

End Sub

Private Sub FormLoadItems()

txtRFNumber.Locked = True
txtGNumber.Locked = True
txtGName.Locked = True
txtRoomType.Locked = True
txtRoomNumber.Locked = True
txtCheckInDate.Locked = True
txtCheckOutDate.Locked = True
txtRoomRate.Locked = True
txtCheckInTime.Locked = True
txtCheckOutTime.Locked = True
txtTotalRoomPayment.Locked = True
txtTotalAmenitiesPayment.Locked = True
txtCheckOutTimePayment.Locked = True
txtTotal.Locked = True
txtChange.Locked = True
xpbtCheckIn.Enabled = False
txtBookTotal.Locked = True
txtBookNextPay.Locked = True
txtBookPay.Locked = True
txtReservationTotal.Locked = True

txtGNumber.Text = frmCheckIn.txtGNumber.Text
txtRoomType.Text = frmCheckIn.xpcboRoomType.Text
txtRoomNumber.Text = frmCheckIn.xpcboRoomNumber.Text
txtCheckInDate.Text = frmCheckIn.DTPicker1.Value
txtCheckOutDate.Text = frmCheckIn.dtpickCheckOut.Value
txtCheckInTime.Text = Format$(Now, "hh:mm AM/PM")
txtCheckOutTime.Text = frmCheckIn.xpcboExtendTime.Text
txtRoomRate.Text = frmCheckIn.txtRoomRate.Text
txtTotalAmenitiesPayment.Text = frmCheckIn.txtTotalAmenities.Text

If frmCheckIn.txtDateDiff.Text = 0 Then
txtTotalRoomPayment.Text = Val(txtRoomRate.Text) * 1

ElseIf frmCheckIn.txtDateDiff > 0 Then
txtTotalRoomPayment.Text = Val(txtRoomRate.Text) * Val(frmCheckIn.txtDateDiff.Text)
Else
MsgBox "Invalid Number Of Days.Please Go Back and Change it", vbCritical + vbOKCancel, "Error Computing Total Room Payment"
End If

strSQL = "SELECT tblRoom_Status.room_num, tblRoom_Status.room_type, " & _
         "tblRoom_Info.room_name, tblRoom_Info.room_rate, " & _
         "tblRoom_Info.room_rate_per_hour, tblRoom_Info.room_rate_on_season, " & _
         "tblRoom_Info.room_description, tblRoom_Status.room_status " & _
         "FROM tblRoom_Info INNER JOIN tblRoom_Status ON " & _
         "tblRoom_Info.room_type_num = tblRoom_Status.room_type " & _
         "WHERE (((tblRoom_Status.room_num)=" & txtRoomNumber.Text & "));"
         
   Set recSet = New ADODB.Recordset
   
    With recSet
    .Open strSQL, Conn, 3, 2
     
    PerHourRate = !room_rate_per_hour
    
    .Close
    
    End With
    
Select Case frmCheckIn.txtCheckOutTimeIndex.Text

Case "-1"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 0

Case "0"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 0

Case "1"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 1

Case "2"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 2

Case "3"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 3

Case "4"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 4

Case "5"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 5

Case "6"

txtCheckOutTimePayment.Text = Val(PerHourRate) * 6

End Select


txtTotal.Text = Val(txtTotalRoomPayment.Text) + Val(txtTotalAmenitiesPayment.Text) + Val(txtCheckOutTimePayment.Text)
txtBookTotal.Text = Val(txtTotal.Text)
txtBookPay.Text = Val(txtBookTotal.Text) / 2
txtReservationTotal.Text = Val(txtTotal.Text)
txtBookNextPay.Text = Val(txtTotal.Text) / 2
End Sub

Private Sub GetCustomerName()

strSQL = "select * from tblcustomer_info where customer_num= " & txtGNumber.Text & ""

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

txtGName.Text = !customer_fname & " " & !customer_lname

.Close

End With

End Sub

Private Sub GetRFNumFromBookingInfo()

If frmCheckIn.txtCmdConfirmTag.Text = "1" Then

strSQL = "select max(record_num) as MaxRFNum from tblbooking_info"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

txtRFNumber.Text = "WI-" & !maxrfnum + 1

.Close

End With

ElseIf frmCheckIn.txtCmdConfirmTag.Text = "2" Then

strSQL = "select max(record_num) as MaxRFNum from tblbooking_info"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

txtRFNumber.Text = "BK-" & !maxrfnum + 1

.Close

End With

txtCheckInTime.Text = Format$("12:00 PM", "hh:mm AM/PM")

ElseIf frmCheckIn.txtCmdConfirmTag.Text = "3" Then

strSQL = "select max(record_num) as MaxRFNum from tblreservation_info"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

txtRFNumber.Text = "RV-" & !maxrfnum + 1

.Close

End With

txtCheckInTime.Text = Format$("12:00 PM", "hh:mm AM/PM")

End If

End Sub

Private Sub txtAmoutPaid_Change()
If Val(txtAmoutPaid.Text) < Val(txtTotal.Text) Then
txtChange.Text = ""
xpbtCheckIn.Enabled = False
ElseIf Val(txtAmoutPaid.Text) >= Val(txtTotal.Text) Then
txtChange.Text = (txtAmoutPaid.Text) - Val(txtTotal.Text)
xpbtCheckIn.Enabled = True
End If
End Sub

Private Sub txtBookPaid_Change()

        If Val(txtBookPaid.Text) >= Val(txtBookPay.Text) Then
        txtBookChange.Text = Val(txtBookPaid.Text) - Val(txtBookPay.Text)
        xpbtCheckIn.Enabled = True
        Else
        txtBookChange.Text = 0
        xpbtCheckIn.Enabled = False
        End If
        
End Sub



Private Sub txtBookPay_Change()
txtBookNextPay.Text = Val(txtBookTotal.Text) - Val(txtBookPay.Text)
If chkboxBookGreaterPay.Value = 1 Then
txtBookPaid.Text = Val(txtBookPay.Text)
End If
End Sub


Private Sub xpbtBack_Click()
Unload Me
End Sub

Private Sub ListViewProp()

With ListView1
 
 .View = lvwReport
 .FullRowSelect = True
 .GridLines = True
 .ColumnHeaders.Clear
 .ColumnHeaders.Add 1, , "Item Name", .Width * 0.34
 .ColumnHeaders.Add 2, , "Item Price", .Width * 0.34
 .ColumnHeaders.Add 3, , "Quantity", .Width * 0.28

End With
End Sub

Private Sub MigrateListItem()

ListCount = frmCheckIn.ListView1.ListItems.Count

For i = 1 To ListCount
Set objList = ListView1.ListItems.Add(, , frmCheckIn.ListView1.ListItems(i).SubItems(1))
    objList.SubItems(1) = frmCheckIn.ListView1.ListItems(i).SubItems(2)
    objList.SubItems(2) = frmCheckIn.ListView1.ListItems(i).SubItems(3)

Next i

End Sub

Private Sub CheckInGuest()

If frmCheckIn.txtCmdConfirmTag.Text = "1" Then
'this add new record on tblbooking_record

strSQL = "select * from tblbooking_record"

Set recSet = New ADODB.Recordset

    With recSet
    .Open strSQL, Conn, 3, 3
    
    .AddNew
    
    !customer_num = txtGNumber.Text
    BookingNumber = !booking_num 'this is needed on tblbooking_info
    
    .Update
    
 '   .Close

    End With
    
    
'this add new record on tblbooking_info

strSQL = "select * from tblbooking_info"

    Set recSet = New ADODB.Recordset
    
    With recSet
    .Open strSQL, Conn, 3, 3
    
    .AddNew
    
    !booking_num = BookingNumber 'the one on above codes
    !booking_date = txtCheckInDate.Text
    !out_date = txtCheckOutDate.Text
    !room_num = txtRoomNumber.Text
    !user_id = frmMain.txtUserID
    !date_done = txtCheckInDate.Text
    !booking_status = 0
    !expected_time_in = frmCheckIn.txtTimeIn.Text
    !expected_time_out = txtCheckOutTime.Text
    !check_in_time = txtCheckInTime.Text
    !tran_type = txtRFNumber.Text
    
    .Update
    
    
'    .Close
    
    End With
    
    
'this will make the selected room occupied
'if your reading this Your Cool! /m/,

strSQL = "select * from tblroom_status where room_num= " & txtRoomNumber.Text & ""

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 2

 !room_status = 1
 !customer_number = txtGNumber.Text
 .Update

'.Close

End With


'this will get the max from tblbooking_info and will be used for billing

strSQL = "select max(record_num) as MaxRecordBooking from tblbooking_info"

Set recSet = New ADODB.Recordset

With recSet

.Open strSQL, Conn, 3, 3

BookingMax = !maxrecordbooking

'.Close

End With


'add new record on billing

strSQL = "select * from tblbilling_booking"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!record_num = BookingMax
!advance_payment = 0
!others = Val(txtTotalAmenitiesPayment.Text) + Val(txtCheckOutTimePayment.Text)
!total = txtTotal.Text
!refund = 0
!user_id = frmMain.txtUserID
!date_done = Format$(Now, "mm/dd/yyyy")
!remaining_balance = 0
!second_payment = 0

.Update

'.Close
End With


'this will get the max on tblbilling and will be used on tblbilling amenities

strSQL = "select max(billing_num) as MaxBillingNum from tblbilling_booking"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3
BillingMax = !maxbillingnum

'.Close
End With


'add new records on tblbilling_amenities,listing all the items that the guest requested

strSQL = "select * from tblbilling_amenities"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

If ListCount > 0 Then

 For a = 1 To ListCount
 .AddNew
!billing_num = BillingMax
!item_num = frmCheckIn.ListView1.ListItems(a)
!item_quantity = frmCheckIn.ListView1.ListItems(a).SubItems(3)
.Update

Next a

'.Close

'Else
' .Close

End If

End With

strSQL = "select * from tbldatesandtime"

Set recSet = New ADODB.Recordset

    With recSet
    .Open strSQL, Conn, 3, 3
    
    .AddNew
    
    !transaction_type = txtRFNumber.Text
    !date_check_in = txtCheckInDate.Text
    !date_check_out = txtCheckOutDate.Text
    !check_in_time = frmCheckIn.txtTimeIn.Text
    !check_out_time = txtCheckOutTime.Text
    !room_number = txtRoomNumber.Text
    
    .Update
    
    .Close
    
    End With
    
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'for booking
    
ElseIf frmCheckIn.txtCmdConfirmTag.Text = "2" Then

strSQL = "select * from tblbooking_record"

Set recSet = New ADODB.Recordset

    With recSet
    .Open strSQL, Conn, 3, 3
    
    .AddNew
    
    !customer_num = txtGNumber.Text
    BookingNumber = !booking_num 'this is needed on tblbooking_info
    
    .Update
    
 '   .Close

    End With
    
    
'this add new record on tblbooking_info

strSQL = "select * from tblbooking_info"

    Set recSet = New ADODB.Recordset
    
    With recSet
    .Open strSQL, Conn, 3, 3
    
    .AddNew
    
    !booking_num = BookingNumber 'the one on above codes
    !booking_date = txtCheckInDate.Text
    !out_date = txtCheckOutDate.Text
    !room_num = txtRoomNumber.Text
    !user_id = frmMain.txtUserID
    !date_done = txtCheckInDate.Text
    !booking_status = 2
    !expected_time_in = frmCheckIn.txtTimeIn.Text
    !expected_time_out = txtCheckOutTime.Text
    !tran_type = txtRFNumber.Text
    
        .Update
    
    
'    .Close
    
    End With
    
'this will get the max from tblbooking_info and will be used for billing

strSQL = "select max(record_num) as MaxRecordBooking from tblbooking_info"

Set recSet = New ADODB.Recordset

With recSet

.Open strSQL, Conn, 3, 3

BookingMax = !maxrecordbooking

'.Close

End With


'add new record on billing

strSQL = "select * from tblbilling_booking"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!record_num = BookingMax
!advance_payment = txtBookPay.Text
!others = Val(txtTotalAmenitiesPayment.Text) + Val(txtCheckOutTimePayment.Text)
!total = txtTotal.Text
!refund = Val(txtBookPay.Text) / 2
!user_id = frmMain.txtUserID
!date_done = Format$(Now, "mm/dd/yyyy")
!remaining_balance = txtBookNextPay.Text
!second_payment = txtBookNextPay.Text

.Update

'.Close
End With


'this will get the max on tblbilling and will be used on tblbilling amenities

strSQL = "select max(billing_num) as MaxBillingNum from tblbilling_booking"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3
BillingMax = !maxbillingnum

'.Close
End With


'add new records on tblbilling_amenities,listing all the items that the guest requested

strSQL = "select * from tblbilling_amenities"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

If ListCount > 0 Then

 For a = 1 To ListCount
 .AddNew
!billing_num = BillingMax
!item_num = frmCheckIn.ListView1.ListItems(a)
!item_quantity = frmCheckIn.ListView1.ListItems(a).SubItems(3)
.Update

Next a

'.Close

'Else
' .Close

End If

End With

strSQL = "select * from tbldatesandtime"

Set recSet = New ADODB.Recordset

    With recSet
    .Open strSQL, Conn, 3, 3
    
    .AddNew
    
    !transaction_type = txtRFNumber.Text
    !date_check_in = txtCheckInDate.Text
    !date_check_out = txtCheckOutDate.Text
    !check_in_time = frmCheckIn.txtTimeIn.Text
    !check_out_time = txtCheckOutTime.Text
    !room_number = txtRoomNumber.Text
    
    .Update
    
    .Close
    
    End With
    
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'this is for reservation
    
ElseIf frmCheckIn.txtCmdConfirmTag.Text = "3" Then


strSQL = "select * from tblreservation_record"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!customer_num = txtGNumber.Text
ReservationNumber = !reservation_num

.Update

End With

strSQL = "select * from tblreservation_info"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!reservation_num = ReservationNumber
!room_num = txtRoomNumber.Text
!reservation_date = txtCheckInDate.Text
!out_date = txtCheckOutDate.Text
!user_id = frmMain.txtUserID.Text
!date_done = Format$(Now, "mm/dd/yyyy")
!reservation_status = 1
!expected_time_in = "2:00 PM"
!expected_time_out = txtCheckOutTime.Text
!total_payment = txtReservationTotal.Text
!tran_type = txtRFNumber.Text

.Update

End With

strSQL = "select * from tbldatesandtime"

Set recSet = New ADODB.Recordset

With recSet
.Open strSQL, Conn, 3, 3

.AddNew

!transaction_type = txtRFNumber.Text
!date_check_in = txtCheckInDate.Text
!date_check_out = txtCheckOutDate.Text
!check_in_time = "2:00 PM"
!check_out_time = txtCheckOutTime.Text
!room_number = txtRoomNumber.Text

.Update

.Close

End With
    
End If
    
End Sub

Private Sub xpbtCheckIn_Click()
If MsgBox("Please Make Sure that you print the OR of the guest before continuing. Do you want to continue?", vbInformation + vbYesNo, "Confirm Check-In") = vbYes Then

CheckInGuest
MsgBox "Writing Records in Database Successfull. Returning in the main form", vbInformation + vbOKOnly, "Check In Successfull"
Unload Me
Unload frmCheckIn
Else

Exit Sub

End If
End Sub

