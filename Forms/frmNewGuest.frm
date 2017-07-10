VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmNewGuest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9780
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewGuest.frx":0000
   ScaleHeight     =   9780
   ScaleWidth      =   7950
   Begin XPControls.XPButton cmdSave 
      Height          =   495
      Left            =   2160
      TabIndex        =   46
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Save"
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
   Begin XPControls.XPCombo cboGStatus 
      Height          =   315
      Left            =   5760
      TabIndex        =   44
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Text            =   "Single"
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
   Begin XPControls.XPCombo cboGSex 
      Height          =   315
      Left            =   5760
      TabIndex        =   43
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Text            =   "Male"
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
      Left            =   1920
      TabIndex        =   29
      Top             =   1335
      Width           =   1815
   End
   Begin VB.TextBox txtFillUpDate 
      Height          =   495
      Left            =   12120
      TabIndex        =   28
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox txtUserID 
      Height          =   495
      Left            =   10680
      TabIndex        =   27
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox txtPContactnum 
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
      Left            =   2640
      TabIndex        =   26
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox txtPAddress 
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
      Left            =   2640
      TabIndex        =   25
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox txtPName 
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
      Left            =   2640
      TabIndex        =   24
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox txtGPurpose 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   23
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtGAge 
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
      Left            =   5760
      TabIndex        =   22
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtGPassportNum 
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox txtGEmail 
      Height          =   285
      Left            =   5760
      TabIndex        =   20
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtGContactNum 
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
      Left            =   5760
      TabIndex        =   19
      Top             =   2055
      Width           =   1815
   End
   Begin VB.TextBox txtGNationality 
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
      Left            =   1920
      TabIndex        =   18
      Top             =   4215
      Width           =   1815
   End
   Begin VB.TextBox txtGAddress 
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
      Left            =   1920
      TabIndex        =   17
      Top             =   3495
      Width           =   1815
   End
   Begin VB.TextBox txtGLname 
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
      Left            =   1920
      TabIndex        =   16
      Top             =   2775
      Width           =   1815
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
      Left            =   1920
      TabIndex        =   15
      Top             =   2055
      Width           =   1815
   End
   Begin XPControls.XPButton cmdClear 
      Height          =   495
      Left            =   3600
      TabIndex        =   47
      Top             =   8640
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
   Begin XPControls.XPButton cmdBack 
      Height          =   495
      Left            =   5040
      TabIndex        =   48
      Top             =   8640
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
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "New Guest Registration"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   45
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   120
      Picture         =   "frmNewGuest.frx":455B
      Top             =   120
      Width           =   705
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   3375
      Left            =   0
      Top             =   6000
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   5055
      Left            =   0
      Top             =   960
      Width           =   7935
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Required Fields"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   0
      TabIndex        =   42
      Top             =   9480
      Width           =   2040
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5520
      TabIndex        =   41
      Top             =   4080
      Width           =   120
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5520
      TabIndex        =   40
      Top             =   3360
      Width           =   120
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5520
      TabIndex        =   39
      Top             =   2640
      Width           =   120
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2400
      TabIndex        =   38
      Top             =   7920
      Width           =   120
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2400
      TabIndex        =   37
      Top             =   7200
      Width           =   120
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2400
      TabIndex        =   36
      Top             =   6480
      Width           =   120
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5520
      TabIndex        =   35
      Top             =   1920
      Width           =   120
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1680
      TabIndex        =   34
      Top             =   4080
      Width           =   120
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1680
      TabIndex        =   33
      Top             =   3360
      Width           =   120
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1680
      TabIndex        =   32
      Top             =   2640
      Width           =   120
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1680
      TabIndex        =   31
      Top             =   1920
      Width           =   120
   End
   Begin VB.Label Label16 
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   30
      Top             =   1200
      Width           =   1560
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      TabIndex        =   14
      Top             =   7920
      Width           =   1800
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      TabIndex        =   13
      Top             =   7200
      Width           =   960
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      TabIndex        =   12
      Top             =   6480
      Width           =   600
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Person to Contact in Case of Emergency"
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
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   4560
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose of Stay:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   1920
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3960
      TabIndex        =   9
      Top             =   4080
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3960
      TabIndex        =   8
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3960
      TabIndex        =   7
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passport Num:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   1560
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3960
      TabIndex        =   5
      Top             =   4800
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Num:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3960
      TabIndex        =   4
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   600
   End
End
Attribute VB_Name = "frmNewGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
clear_Fields
End Sub

Private Sub cmdClear1_Click()

End Sub

Private Sub cmdSave_Click()
 If txtGName.Text <> "" And txtGLname.Text <> "" And txtGAddress.Text <> "" And txtGNationality.Text <> "" And _
  txtGContactNum.Text <> "" And txtGAge.Text <> "" And cboGSex.Text <> "" And cboGStatus.Text <> "" _
   And txtPName.Text <> "" And txtPAddress.Text <> "" And txtPContactnum.Text <> "" Then '0
   
     If MsgBox("Are You Sure You Want To Save This Guest Information?", vbInformation + vbYesNo, "Confirm Save") = vbYes Then '1
     
       strSQL = "select * from tblcustomer_info"
       
       Set recSet = New ADODB.Recordset
       
       With recSet
       
        .Open strSQL, Conn, 3, 3
        .AddNew
        
        !customer_fname = StrConv(txtGName.Text, vbProperCase)
        !customer_lname = StrConv(txtGLname.Text, vbProperCase)
        !customer_address = StrConv(txtGAddress.Text, vbProperCase)
        !customer_nationality = StrConv(txtGNationality.Text, vbProperCase)
        !customer_contact_num = StrConv(txtGContactNum.Text, vbProperCase)
        !customer_email = txtGEmail.Text
        !customer_passport_num = txtGPassportNum.Text
        !customer_age = txtGAge.Text
        !customer_sex = cboGSex.Text
        !customer_status = cboGStatus.Text
        !purpose_of_stay = txtGPurpose.Text
        !person_to_contact_in_case_of_emergency = StrConv(txtPName.Text, vbProperCase)
        !person_address = StrConv(txtPAddress.Text, vbProperCase)
        !person_contact_num = txtPContactnum.Text
        !fill_up_date = txtFillUpDate.Text
        !user_id = txtUserID.Text
        
        .Update
        .Close
        
        If MsgBox("Successfully Saved!If you want to book/reserve with this guest info press YES", vbInformation + vbYesNo, "Success") = vbYes Then '2
          
          frmChoose.txtGuestNumber = txtGNumber.Text 'this will be needed when the user choose walk-in
          
          Unload Me
          
          frmChoose.Show vbModal, frmMain
          

          

          
        Else '2
        
        MsgBox "Returning to Main Form..", vbOKOnly, "Returning"
        
        Unload Me
        
        End If '2
       
       End With
     
     Else '1
     
       If MsgBox("Saving Cancelled!Do you want to return in the Main Form?", vbInformation + vbYesNo, "Cancel") = vbYes Then '3
       
        MsgBox "Returning to Main Form...", vbOKOnly, "Returning"
        
        Unload Me
        
       Else
       
       
       End If '3
     
     End If '1
     
 Else '0
 
 MsgBox "Please Fill All Required Fields.", vbInformation + vbOKOnly, "Error"
 
 End If '0
End Sub

Private Sub Form_Load()
FormPos frmNewGuest
Me.KeyPreview = True

DBCON

cbo_GuestItems
generate_GuestNo

txtGNumber.Locked = True
txtGNumber.ForeColor = vbRed
txtFillUpDate.Enabled = False
txtUserID.Enabled = False

txtUserID.Text = frmMain.txtUserID.Text
txtFillUpDate.Text = Date
End Sub

Sub clear_Fields()

txtGName.Text = ""
txtGLname.Text = ""
txtGAddress.Text = ""
txtGNationality.Text = ""
txtGContactNum.Text = ""
txtGEmail.Text = ""
txtGAge.Text = ""
txtGPassportNum.Text = ""
txtGPurpose.Text = ""
txtPName.Text = ""
txtPAddress.Text = ""
txtPContactnum.Text = ""
txtGName.SetFocus

End Sub

Sub generate_GuestNo()

 strSQL = "select max(customer_num) as maxGuest_Num from tblcustomer_info"
 
  Set recSet = New ADODB.Recordset
  
   With recSet
   
    .Open strSQL, Conn, 3, 3
    
    If .BOF Or IsNull(!maxguest_num) Then
    
     txtGNumber.Text = "1"
     
    Else
    
     txtGNumber.Text = !maxguest_num + 1
    
    End If
    
    .Close
    
   End With

End Sub

Sub cbo_GuestItems()

 With cboGSex
  .AddItem "Male"
  .AddItem "Female"
 End With
 
 With cboGStatus
 .AddItem "Single"
 .AddItem "Married"
 End With
End Sub



Private Sub txtGAge_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtGAge) 'uses the module modNumericOnly

End Sub

Private Sub txtGContactNum_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtGContactNum) 'uses the module modNumericOnly

End Sub


Private Sub txtPContactnum_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtPContactnum) 'uses the module modNumericOnly

End Sub

