VERSION 5.00
Begin VB.Form frmUpdateRoom 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7095
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   13410
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRoomTypeNum 
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   3840
      Width           =   10935
   End
   Begin VB.TextBox txtMaxPax 
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtRateHour 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtSeasonRate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtRoomRate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboRoomType 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "*all fields are required"
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Description:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Max Pax:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Rate Per Hour:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Seasonal Rate:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Room Rate:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Room Type:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmUpdateRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRoomType_Click()
what_RoomChoose
End Sub

Sub what_RoomChoose()
' this helps to identify what to edit in database, it is based on room_type_num
Select Case cboRoomType.ListIndex

Case 0

txtRoomTypeNum.Text = 1 ' if fan room

Case 1

txtRoomTypeNum.Text = 2 ' if standard room

Case 2

txtRoomTypeNum.Text = 3 'if deluxe room

Case 3

txtRoomTypeNum.Text = 4 ' if family room

Case 4

txtRoomTypeNum.Text = 5 'if suite room


End Select

End Sub


Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
clear_Fields
End Sub

Private Sub cmdSave_Click()
If txtRoomRate.Text <> "" And txtSeasonRate.Text <> "" And txtRateHour.Text <> "" And txtMaxPax.Text <> "" _
   And txtDescription <> "" And cboRoomType.Text <> "" Then '0
   
    If MsgBox("Are You sure You Want save changes on the current record?", vbInformation + vbYesNo, "Confirm Save") = vbYes Then '1
   
        strSQL = "select * from tblroom_info where room_type_num=" & txtRoomTypeNum.Text & ""
 
        Set recSet = New ADODB.Recordset
 
            With recSet
 
                 .Open strSQL, Conn, 3, 3
                 
                   !room_rate = txtRoomRate.Text
                   !room_rate_on_season = txtSeasonRate.Text
                   !room_rate_per_hour = txtRateHour.Text
                   !room_max_pax = txtMaxPax.Text
                   !room_description = txtDescription.Text
                   
                   .Update
  
                MsgBox "Save Successfull! Returning in the Room Information Form", vbInformation + vbOKOnly, "Success"
                
                Unload Me
  
                 .Close
 
            End With
            
     Else '1
            
       MsgBox "Returning in to Main Form", vbInformation + vbOKOnly, "Save Cancelled"
       
       Unload Me
       
       Unload frmRoomMaintenance
            
     End If '1
   
   
Else '0

 MsgBox "Please Fill All Fields!", vbExclamation + vbOKOnly, "Missing Text"
   
End If '0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyBack Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
DBCON
cboItems
clear_Fields
txtRoomTypeNum.Locked = True
End Sub

Sub clear_Fields()

txtRoomRate.Text = ""
txtSeasonRate.Text = ""
txtRateHour.Text = ""
txtMaxPax.Text = ""
txtDescription.Text = ""

End Sub

Sub cboItems()
With cboRoomType
.AddItem "Fan Room", 0
.AddItem "Standard Room", 1
.AddItem "Deluxe Room", 2
.AddItem "Family Room", 3
.AddItem "Suite Room", 4
End With
End Sub

Private Sub txtMaxPax_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtMaxPax)

End Sub

Private Sub txtRateHour_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtRateHour)

End Sub

Private Sub txtRoomRate_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtRoomRate) 'uses the module modNumericOnly

End Sub


Private Sub txtSeasonRate_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtSeasonRate)

End Sub


