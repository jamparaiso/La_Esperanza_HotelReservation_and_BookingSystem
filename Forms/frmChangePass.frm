VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmChangePass 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5460
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5055
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChangePass.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPButton cmdChange 
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Ch&ange Password"
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
   Begin VB.TextBox txtConNewPWord 
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtNewPWord 
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
      Left            =   2880
      TabIndex        =   5
      Top             =   2775
      Width           =   1695
   End
   Begin VB.TextBox txtPWord 
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2055
      Width           =   1695
   End
   Begin VB.TextBox txtUname 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1335
      Width           =   1695
   End
   Begin XPControls.XPButton cmdBack 
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4200
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
   Begin XPControls.XPButton cmdClear 
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   4200
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   """ and ' is disabled please see help for details"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Change User Password"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   3360
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   120
      Picture         =   "frmChangePass.frx":455B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm New Password:"
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
      Top             =   3480
      Width           =   2520
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
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
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password:"
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
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Width           =   1080
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdChange_Click()
 If Not txtNewPWord.Text = txtConNewPWord.Text Then ' if the two new passwords didnt match

     MsgBox "Password didn't match! Try Again", vbInformation + vbOKOnly, "Error"
     
     txtNewPWord.Text = ""
     txtConNewPWord.Text = ""
     txtNewPWord.SetFocus

 Else

     validate_userAcct 'if the two new passwords match,the system will proceed to this function

 End If

End Sub

Private Sub cmdClear_Click()
clear_Fields
End Sub

Private Sub Form_Load()
FormPos frmChangePass
Me.KeyPreview = True
DBCON
End Sub

Sub validate_userAcct()

strSQL = "select uname,pword from tbluser_info " & _
"where uname= '" & txtUname.Text & "' and pword= '" & txtPWord.Text & " ' "

Set recSet = New ADODB.Recordset

 With recSet
  .Open strSQL, Conn, 3, 2
  
    If .EOF Or .BOF Then '1 no records found,null
    
      If MsgBox("Invalid Username or Password!Try Again?", vbCritical + vbYesNo, "Error") = vbYes Then '2
      
       clear_Fields
      
      Else
      
       MsgBox "Returning to Main Form", vbInformation + vbOKOnly, "Cancel"
      
       Unload Me
       
      End If '2
      
    Else 'the tbl is not empty and found the user
    
        If MsgBox("Are You Sure You Want To Change Your Current Password?", vbInformation + vbYesNo, "Confirm Change Password") = vbYes Then '3
    
                 !pword = txtConNewPWord.Text 'successfully changed
                 .Update
                 
                 MsgBox "Changing Password Successful! Returning to Main Form...", vbInformation + vbOKOnly, "Success"
                 
                 Unload Me
                 
        Else
        
                MsgBox "Returning to Main Form", vbInformation + vbOKOnly, "Change Password Cancel"
     
        End If '3
    
    End If '1
    
  .Close
  
 End With
 
End Sub

Sub clear_Fields()
txtUname.Text = ""
txtPWord.Text = ""
txtNewPWord.Text = ""
txtConNewPWord.Text = ""
txtUname.SetFocus
End Sub


Private Sub txtConNewPWord_Change()

End Sub

Private Sub txtConNewPWord_KeyPress(KeyAscii As Integer)
Select Case KeyAscii

Case 34, 39

KeyAscii = 0

End Select
End Sub

Private Sub txtNewPWord_KeyPress(KeyAscii As Integer)
Select Case KeyAscii

Case 34, 39

KeyAscii = 0

End Select
End Sub
