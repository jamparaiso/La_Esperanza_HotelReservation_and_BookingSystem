VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmChoose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Transaction Type"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoose.frx":0000
   ScaleHeight     =   2130
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPCombo cboChooseType 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Text            =   "Choose Here"
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
   Begin VB.TextBox txtGuestNumber 
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin XPControls.XPButton XPButton1 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Ok"
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
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   0
      Picture         =   "frmChoose.frx":455B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this form is shown when the user finished guest registration, so that he/she can choose what transaction is be made

Sub cboChoose_Item()
 With cboChooseType
 .AddItem "Booking", 0
 .AddItem "Reservation", 1
 .AddItem "Check-In", 2
 End With
End Sub



Private Sub Form_Load()
FormPos frmChoose
cboChoose_Item

End Sub



Private Sub XPButton1_Click()
If cboChooseType.listIndex = 0 Then 'for booking

frmCheckIn.txtGnumber.Text = txtGuestNumber.Text

frmMain.mnuNewBookWalk_Click


Unload Me

ElseIf cboChooseType.listIndex = 1 Then

frmCheckIn.txtGnumber.Text = txtGuestNumber.Text

frmMain.mnuNewReservation_Click


Unload Me

ElseIf cboChooseType.listIndex = 2 Then 'for walk-in

frmCheckIn.txtGnumber.Text = txtGuestNumber.Text

frmMain.ForWalkIn


Unload Me

Else

MsgBox "Please choose on the choices before continuing.", vbOKOnly, "No Choosen item"

cboChooseType.SetFocus


End If
End Sub

