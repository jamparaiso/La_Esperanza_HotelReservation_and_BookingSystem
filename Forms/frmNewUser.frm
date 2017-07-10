VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmNewUser 
   Caption         =   "New User"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewUser.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPButton cmdBack 
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   6360
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
   Begin XPControls.XPButton cmdClear 
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Clear Fields"
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
   Begin XPControls.XPButton cmdSave 
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   6360
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
   Begin VB.TextBox txtListIndex 
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cboAccountType 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox txtPassword 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   10
      ToolTipText     =   """ and ' is disabled"
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox txtUsername 
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      ToolTipText     =   """ and ' is disabled"
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtMidname 
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtLname 
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtFname 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NEW USER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   5040
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   120
      Picture         =   "frmNewUser.frx":455B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* All fields are required"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   240
      TabIndex        =   14
      Top             =   6000
      Width           =   2625
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "* "" and ' is cannot be used in username and      password"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   0
      TabIndex        =   13
      Top             =   5520
      Width           =   5100
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
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
      Top             =   4920
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
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
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name:"
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
      Top             =   2760
      Width           =   1440
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
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   2040
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
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   600
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormPos frmNewUser
Me.KeyPreview = True
DBCON
Acct_Items
txtListIndex.Visible = True
End Sub

Private Sub cboAccountType_Click()
cboAcct_ListIndex
End Sub

Sub cboAcct_ListIndex()
'this function is needed on save function

Select Case cboAccountType.listIndex

Case 0
txtListIndex.Text = 1 'if owner

Case 1
txtListIndex.Text = 2 ' if manager

Case 2
txtListIndex.Text = 3 ' if receptionist

Case 3
txtListIndex.Text = 4 'if admin

End Select

End Sub

Private Sub cmdSave_Click()
If txtFname.Text <> "" And txtLname.Text <> "" And txtMidname.Text <> "" And txtUsername.Text <> "" _
And txtPassword.Text <> "" And cboAccountType.Text <> "" Then '0 'if the text fields have blanks

    strSQL = "SELECT tblUser_Info.user_id, tblUser_Info.first_name, tblUser_Info.last_name, tblUser_Info.mid_name, " & _
             "tblUser_Info.uname, tblUser_Info.pword, tblUser_Info.acc_type " & _
             "From tblUser_Info " & _
             "WHERE (((tblUser_Info.uname)='" & txtUsername.Text & "') AND ((tblUser_Info.pword)='" & txtPassword.Text & "'));"

    
    Set recSet = New ADODB.Recordset
    
    With recSet
    
    .Open strSQL, Conn, 3, 2
    
     If .EOF Or .BOF Then '1 tbl is empty, no user duplication
      
      If MsgBox("Are You Sure You Want To Save This User Info?", vbInformation + vbYesNo, "Confirm Save") = vbYes Then '2
       
         .AddNew
            !first_name = txtFname.Text
            !last_name = txtLname.Text
            !mid_name = txtMidname.Text
            !uname = txtUsername.Text
            !pword = txtPassword.Text
            !acc_type = txtListIndex.Text
         .Update
         .Close
         
        MsgBox "Successfully Saved!The System will return on the Main Form.", vbInformation + vbOKOnly, _
                        "Saved"
        
        Unload Me
        
        Else 'save aborted, no records saved
        
            MsgBox "Returning To Main Form", vbInformation + vbOKOnly, "Save Cancel"
    
            Unload Me
         
       End If '2
     
     Else 'tbl not null refer to if 1
     
     MsgBox "Username is already taken!Please choose another one.", vbCritical + vbOKOnly, "Existing Username"
     
     txtUsername.Text = ""
     txtUsername.SetFocus
     
     End If '1
    
    End With

    
Else 'if there are blank fields

MsgBox "Please Fill All Fields First!", vbInformation + vbOKOnly, "Missing Text"

End If '0
End Sub

Sub clear_Fields()
txtFname.Text = ""
txtLname.Text = ""
txtMidname.Text = ""
txtUsername.Text = ""
txtPassword.Text = ""
End Sub

Sub Acct_Items()
'cboaccounttype items

    With cboAccountType
    
        .AddItem "Owner", 0
        .AddItem "Manager", 1
        .AddItem "Receptionist", 2
        .AddItem "Admin", 3
    
    End With
End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
clear_Fields
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Or KeyAscii = 39 Then
KeyAscii = 0

txtUsername.ToolTipText = "These Characters are not allowed:' , """
End If
End Sub


Private Sub txtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Or KeyAscii = 39 Then
KeyAscii = 0

txtUsername.ToolTipText = "These Characters are not allowed:' , """
End If
End Sub
